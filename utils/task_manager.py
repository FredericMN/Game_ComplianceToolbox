import sys
import threading
import atexit

# 任务管理器类定义
class TaskManager:
    """管理爬虫和版号匹配任务，确保在应用关闭时能够正确终止所有任务"""
    
    def __init__(self):
        self.thread_pools = []  # 存储所有活动的线程池
        self.webdrivers = []    # 存储所有活动的WebDriver实例
        self.tasks_running = False
        self.lock = threading.Lock()  # 用于同步访问共享资源
        
        # 注册应用退出时的清理函数
        atexit.register(self.cleanup_all_resources)
    
    def register_thread_pool(self, pool):
        """注册线程池以便在退出时关闭"""
        with self.lock:
            self.thread_pools.append(pool)
            self.tasks_running = True
    
    def unregister_thread_pool(self, pool):
        """取消注册线程池"""
        with self.lock:
            if pool in self.thread_pools:
                self.thread_pools.remove(pool)
            if not self.thread_pools:
                self.tasks_running = False
    
    def register_webdriver(self, driver):
        """注册WebDriver实例以便在退出时关闭"""
        with self.lock:
            self.webdrivers.append(driver)
    
    def unregister_webdriver(self, driver):
        """取消注册WebDriver实例"""
        with self.lock:
            if driver in self.webdrivers:
                self.webdrivers.remove(driver)
    
    def cleanup_all_resources(self):
        """清理所有资源，包括线程池和WebDriver实例"""
        print("开始清理所有爬虫任务资源...")
        
        active_pools = []
        active_drivers = []

        # --- Shutdown Thread Pools First ---
        with self.lock:
            # Make copies to avoid modification issues during iteration
            active_pools = self.thread_pools[:] 
            self.thread_pools.clear() # Clear original list
        
        print(f"尝试关闭 {len(active_pools)} 个线程池...")
        for pool in active_pools:
            try:
                # Shutdown non-blocking first to signal cancellation
                pool.shutdown(wait=False, cancel_futures=True) 
            except Exception as e:
                print(f"关闭线程池 (阶段1) 时出错: {str(e)}")
        # Optionally add a small delay or a wait phase here if needed
        print("线程池关闭信号已发送。")


        # --- Quit WebDriver Instances ---
        with self.lock:
             # Make copies
            active_drivers = self.webdrivers[:]
            self.webdrivers.clear() # Clear original list

        print(f"尝试关闭 {len(active_drivers)} 个WebDriver实例...")
        for driver in active_drivers:
            print(f"  正在关闭 WebDriver: {driver}")
            try:
                driver.quit()
                print(f"  WebDriver 关闭成功: {driver}")
            except Exception as e:
                # Log error but continue cleanup
                print(f"  关闭WebDriver时出错: {driver}, 错误: {str(e)}") 
        print("所有活动的WebDriver实例已尝试关闭。")


        # --- Kill Lingering Driver Processes ---
        print("开始清理 msedgedriver 进程...")
        # 不再尝试从 main 导入，直接使用 psutil 清理
        try:
            import psutil
            terminated_count = 0
            failed_count = 0
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    p_name = proc.info.get('name') # 使用 get 避免 KeyError
                    p_pid = proc.info.get('pid')
                    if p_name and p_name.lower() == 'msedgedriver.exe':
                        print(f"  找到 msedgedriver 进程 (PID: {p_pid}), 尝试终止...")
                        try:
                            # 先尝试温和终止
                            proc.terminate()
                            # 等待最多2秒
                            gone, alive = psutil.wait_procs([proc], timeout=2)
                            if proc in alive:
                                print(f"    进程 (PID: {p_pid}) 未在2秒内终止，强制终止...")
                                proc.kill()
                            print(f"    进程 (PID: {p_pid}) 已终止。")
                            terminated_count += 1
                        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                            # 进程可能已经消失或无权限访问
                            print(f"    终止进程 (PID: {p_pid}) 时出错或进程已消失。")
                            failed_count += 1
                        except Exception as term_err:
                             print(f"    终止进程 (PID: {p_pid}) 时出现意外错误: {term_err}")
                             failed_count += 1

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass # Process already gone or inaccessible during iteration
                except Exception as proc_err:
                     print(f"    处理进程信息时出错 (PID: {proc.info.get('pid', 'N/A')}): {proc_err}")
                     failed_count += 1 # 计入失败

            if terminated_count > 0 or failed_count > 0:
                 print(f"驱动进程清理：已终止 {terminated_count} 个，失败 {failed_count} 个。")
            else:
                 print("驱动进程清理：未找到活动的 msedgedriver 进程。")

        except ImportError:
            print("错误：无法导入 psutil 库，无法执行 msedgedriver 进程清理。请确保已安装 psutil。")
        except Exception as e:
            print(f"执行 msedgedriver 进程清理时出错: {str(e)}")
        
        with self.lock:        
            self.tasks_running = False # Mark tasks as no longer running
        print("所有爬虫任务资源清理过程结束。")

# 创建全局任务管理器实例
task_manager = TaskManager() 