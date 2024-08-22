import os
import ctypes
import pyautogui
import asyncio
import psutil
import time
import random
import logging
import codecs
from logging.handlers import RotatingFileHandler
import signal
from typing import Optional, List
import win32gui
import win32con
import win32process
from pywinauto import Application, findwindows
from pywinauto.findwindows import ElementNotFoundError
from config_manager import ConfigManager, ConfigKeys

# 1. Logging Setup
class Logger:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.initialize_logging()

    def initialize_logging(self):
        log_file_path = self.config_manager.get(ConfigKeys.LOG_FILE_PATH, 'afk_script.log')
        max_bytes = self.config_manager.get(ConfigKeys.MAX_BYTES, 5242880)
        backup_count = self.config_manager.get(ConfigKeys.BACKUP_COUNT, 5)

        handler = RotatingFileHandler(log_file_path, maxBytes=max_bytes, backupCount=backup_count)
        handler.setFormatter(logging.Formatter("%(asctime)s:%(levelname)s:%(message)s"))
        handler.setLevel(logging.INFO)

        handler.setStream(codecs.open(log_file_path, 'a', 'utf-8'))

        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)
        Logger.log_info(f"Log file path: {log_file_path}")
        Logger.log_info("Logging initialized")

    @staticmethod
    def log_and_handle_exception(e: Exception, message: str) -> None:
        Logger.log_exception(e)
        Logger.log_info(message)

    @staticmethod
    def log_exception(e: Exception):
        logging.error(f"Exception occurred: {e}")
        logging.error("Traceback:", exc_info=True)
        print(f"Error: {e}")

    @staticmethod
    def log_info(message: str):
        logging.info(message)
        print(message)


# 2. Window Management
class RobloxWindowManager:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.width = config_manager.get(ConfigKeys.WINDOW_WIDTH, 816)
        self.height = config_manager.get(ConfigKeys.WINDOW_HEIGHT, 638)
        self.screen_width, self.screen_height = pyautogui.size()
        self.cached_positions = None  

    def allow_set_foreground_window(self):
        try:
            ctypes.windll.user32.AllowSetForegroundWindow(ctypes.windll.kernel32.GetCurrentProcessId())
        except Exception as e:
            Logger.log_exception(e)

    def ensure_window_visible(self, hwnd):
        rect = win32gui.GetWindowRect(hwnd)
        if rect[0] < 0 or rect[1] < 0:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, rect[2] - rect[0], rect[3] - rect[1], win32con.SWP_SHOWWINDOW)

    def find_roblox_windows(self):
        try:
            windows = findwindows.find_windows(title="Roblox", class_name="WINDOWSCLIENT")
            if not windows:
                Logger.log_info("No Roblox windows found.")
                return None

            if len(windows) > 1:
                Logger.log_info(f"Multiple Roblox windows found: {windows}")

            return windows
        except ElementNotFoundError as e:
            Logger.log_exception(e)
            return None
        except Exception as e:
            Logger.log_exception(e)
            return None

    def bring_window_to_front(self, handle):
        try:
            self.allow_set_foreground_window()

            # Lấy ID của luồng hiện tại và luồng của cửa sổ mục tiêu
            foreground_thread = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[0]
            target_thread = win32process.GetWindowThreadProcessId(handle)[0]

            # Gắn kết các luồng
            if foreground_thread != target_thread:
                ctypes.windll.user32.AttachThreadInput(foreground_thread, target_thread, True)

            win32gui.SetForegroundWindow(handle)
            win32gui.BringWindowToTop(handle)
            time.sleep(0.5)

            # Tách các luồng sau khi hoàn thành
            if foreground_thread != target_thread:
                ctypes.windll.user32.AttachThreadInput(foreground_thread, target_thread, False)

        except win32gui.error as e:
            Logger.log_exception(e)
        except Exception as e:
            Logger.log_exception(e)

    def ensure_window_active(self, handle):
        try:
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
            self.bring_window_to_front(handle)
        except win32gui.error as e:
            Logger.log_exception(e)
        except Exception as e:
            Logger.log_exception(e)

    def position_windows_in_grid(self, handles: List[int], windows_per_batch: int) -> None:
        if self.cached_positions is None:  # Calculate positions only once
            num_windows = windows_per_batch  # Use the batch size to determine grid layout
            cols = min(windows_per_batch, num_windows)
            rows = (num_windows + cols - 1) // cols  # Ceiling division to ensure enough rows

            # Calculate spacing between windows
            x_offset = (self.screen_width - (self.width * cols)) // (cols + 1)
            y_offset = (self.screen_height - (self.height * rows)) // (rows + 1)

            # Precompute positions for a single batch and cache them
            self.cached_positions = []
            for i in range(windows_per_batch):
                row = i // cols
                col = i % cols
                x = col * (self.width + x_offset) + x_offset
                y = row * (self.height + y_offset) + y_offset
                self.cached_positions.append((x, y))

        # Assign positions to windows, reusing cached positions
        for i, handle in enumerate(handles):
            position_index = i % windows_per_batch
            x, y = self.cached_positions[position_index]
            Logger.log_info(f"Processing window {i + 1} with handle {handle} at position ({x}, {y})")
            self.restore_and_resize_window(handle, x, y)

    def ensure_window_restored(self, handle: int) -> None:
        """Ensure that the window is restored from minimized or maximized state."""
        if win32gui.IsIconic(handle):
            win32gui.ShowWindow(handle, win32con.SW_RESTORE)
            time.sleep(1)

        if win32gui.GetWindowPlacement(handle)[1] == win32con.SW_MAXIMIZE:
            win32gui.ShowWindow(handle, win32con.SW_RESTORE)
            time.sleep(1)

    def restore_and_resize_window(self, handle: int, x: int, y: int) -> None:
        try:
            Logger.log_info(f"Attempting to resize window with handle {handle}...")
            
            self.ensure_window_visible(handle)
            self.ensure_window_restored(handle)
            
            win32gui.SetWindowPos(handle, win32con.HWND_TOP, x, y, self.width, self.height, win32con.SWP_SHOWWINDOW)
            Logger.log_info(f"Window with handle {handle} resized to {self.width}x{self.height} at position ({x}, {y}).")
        except win32gui.error as e:
            Logger.log_exception(e)
        except Exception as e:
            Logger.log_exception(e)

    def minimize_window(self, window):
        try:
            window.minimize()
        except Exception as e:
            Logger.log_exception(e)


# 3. Mouse Actions
class MouseActionHandler:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.click_wait_time = config_manager.get(ConfigKeys.CLICK_WAIT_TIME, 1)
        self.taskbar_height = config_manager.get(ConfigKeys.TASKBAR_HEIGHT, 70)
        self.screen_width, self.screen_height = pyautogui.size()

    def get_taskbar_height(self):
        try:
            taskbar_hwnd = win32gui.FindWindow("Shell_TrayWnd", None)
            rect = win32gui.GetWindowRect(taskbar_hwnd)
            taskbar_height = rect[3] - rect[1]
            return taskbar_height
        except Exception as e:
            Logger.log_exception(e)
            return self.taskbar_height

    async def click_random_point_outside(self, window):
        rect = window.rectangle()
        outside_x, outside_y = self.find_random_outside_point(self.screen_width, self.screen_height, rect)
        try:
            pyautogui.moveTo(outside_x, outside_y)
            await asyncio.sleep(self.click_wait_time)
            pyautogui.click()
        except pyautogui.FailSafeException as e:
            Logger.log_exception(e)
        except Exception as e:
            Logger.log_exception(e)

    def find_random_outside_point(self, screen_width, screen_height, rect):
        while True:
            outside_x = random.randint(451, screen_width - 1)
            outside_y = random.randint(0, screen_height - 1)
            if not (rect.left <= outside_x <= rect.right and rect.top <= outside_y <= rect.bottom):
                if outside_y < screen_height - self.taskbar_height:
                    return outside_x, outside_y

    async def click_specific_item(self, window, rel_x=148, rel_y=387):
        rect = window.rectangle()
        abs_x = rect.left + rel_x
        abs_y = rect.top + rel_y
        try:
            pyautogui.moveTo(abs_x, abs_y)
            await asyncio.sleep(self.click_wait_time)
            pyautogui.mouseDown()
            for _ in range(5):
                pyautogui.moveRel(random.randint(-2, 2), random.randint(-2, 2), duration=0.1)
            pyautogui.mouseUp()
            await asyncio.sleep(self.click_wait_time)
            pyautogui.click()
        except pyautogui.FailSafeException as e:
            Logger.log_exception(e)
        except Exception as e:
            Logger.log_exception(e)


# 4. System Controls
class SystemController:
    @staticmethod
    def turn_off_screen():
        try:
            ctypes.windll.user32.SendMessageW(0xFFFF, 0x0112, 0xF170, 2)
            Logger.log_info("Screen turned off.")
        except Exception as e:
            Logger.log_exception(e)


# 5. AFK Bot Core Logic
class RobloxAFKBot:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.window_manager = RobloxWindowManager(config_manager)
        self.mouse_handler = MouseActionHandler(config_manager)
        self.system_controller = SystemController()
        self.active_windows = []
        self.shutdown_flag = False
        Logger(config_manager)
        Logger.log_info("RobloxAFKBot initialized.")

    async def process_window(self, handle: int, retry_count: int = 3) -> None:
        try:
            if not self.window_is_still_open(handle):
                Logger.log_info(f"Roblox window with handle {handle} has been closed unexpectedly.")
                self.active_windows.remove(handle)
                return

            success = False
            for attempt in range(retry_count):
                try:
                    self.window_manager.ensure_window_active(handle)

                    app = Application().connect(handle=handle)
                    window = app.top_window()
                    window.set_focus()

                    await asyncio.sleep(0.1)

                    await self.mouse_handler.click_random_point_outside(window)
                    await asyncio.sleep(0.1)

                    if not self.window_is_still_open(handle):
                        Logger.log_info(f"Window {handle} lost focus after click_random_point_outside. Attempt {attempt + 1}")
                        continue

                    await self.mouse_handler.click_specific_item(window)
                    await asyncio.sleep(0.1)

                    if not self.window_is_still_open(handle):
                        Logger.log_info(f"Window {handle} lost focus after click_specific_item. Attempt {attempt + 1}")
                        continue

                    self.window_manager.minimize_window(window)
                    await asyncio.sleep(0.1)
                    success = True
                    break  # Break the loop if successful
                except ElementNotFoundError as e:
                    Logger.log_info(f"Element not found for window {handle}. Skipping this window.")
                    break  # Skip to the next window if the element is not found
                except Exception as e:
                    Logger.log_exception(e)
                    if attempt < retry_count - 1:
                        Logger.log_info(f"Retrying process_window for handle {handle} (Attempt {attempt + 2})")
                        await asyncio.sleep(0.5 * (2 ** attempt))  # Incremental delay before retrying
                    else:
                        Logger.log_info(f"Failed to process window with handle {handle} after {retry_count} attempts.")

            if not success:
                Logger.log_info(f"Failed to process window with handle {handle}. Moving on to the next window.")
        except Exception as e:
            Logger.log_and_handle_exception(e, f"Critical failure processing window with handle {handle}. Continuing with next window.")


    async def finalize_windows(self):
        try:
            if self.active_windows:
                app = Application().connect(handle=self.active_windows[-1])
                window = app.top_window()
                await self.mouse_handler.click_random_point_outside(window)
                await asyncio.sleep(2)
        except Exception as e:
            Logger.log_exception(e)

    async def process_batches(self) -> None:
        windows_per_batch = self.config_manager.get(ConfigKeys.WINDOWS_PER_BATCH, 1)
        for batch_start in range(0, len(self.active_windows), windows_per_batch):
            batch_handles = self.active_windows[batch_start:batch_start + windows_per_batch]
            self.window_manager.position_windows_in_grid(batch_handles, windows_per_batch=len(batch_handles))
            await self.process_each_window_in_batch(batch_handles, batch_start, windows_per_batch)

    async def process_each_window_in_batch(self, batch_handles: List[int], batch_start: int, windows_per_batch: int) -> None:
        for i, handle in enumerate(batch_handles):
            Logger.log_info(f"[Batch {batch_start // windows_per_batch + 1}] Processing window {i + 1} with handle {handle})")
            await self.process_window(handle)

    async def reset_afk_timer(self) -> None:
        if self.shutdown_flag:
            Logger.log_info("Shutdown flag detected. Exiting AFK timer reset cycle.")
            return

        start_time = time.time()
        Logger.log_info("Starting AFK timer reset cycle.")
        try:
            self.active_windows = self.window_manager.find_roblox_windows()
            if not self.active_windows:
                Logger.log_info("No Roblox windows found. Skipping this cycle.")
                return

            await self.process_batches()

            if not self.shutdown_flag:
                await self.finalize_windows()
        except ElementNotFoundError as e:
            Logger.log_exception(e)
        except Exception as e:
            Logger.log_exception(e)
        finally:
            elapsed_time = time.time() - start_time
            Logger.log_info(f"AFK timer reset cycle completed in {elapsed_time:.2f} seconds.")


    def window_is_still_open(self, handle):
        try:
            return win32gui.GetWindowText(handle) != ''
        except Exception as e:
            Logger.log_exception(e)
            return False

    async def monitor_resources(self):
        while not self.shutdown_flag:
            cpu_usage = psutil.cpu_percent(interval=1)
            memory_usage = psutil.virtual_memory().percent
            Logger.log_info(f"CPU usage: {cpu_usage}%, Memory usage: {memory_usage}%")
            if cpu_usage > 90 or memory_usage > 99:
                Logger.log_info("High resource usage detected. Initiating graceful shutdown.")
                await asyncio.sleep(30)
            else:
                await asyncio.sleep(1800)

    async def main_loop(self):
        while not self.shutdown_flag:
            try:
                self.config_manager.dynamic_reload_config()
                await self.reset_afk_timer()
                if self.shutdown_flag:
                    Logger.log_info("Shutdown flag detected. Exiting script.")
                    break
                sleep_time = random.randint(900, 960)
                Logger.log_info(f"Sleeping for {sleep_time} seconds.")
                for _ in range(sleep_time):
                    if self.shutdown_flag:
                        Logger.log_info("Shutdown flag detected during sleep. Exiting script.")
                        break
                    await asyncio.sleep(1)
            except Exception as e:
                Logger.log_exception(e)
                await asyncio.sleep(60)
        Logger.log_info("RobloxAFKBot has been gracefully shut down.")

    async def run(self):
        Logger.log_info("Script started")
        try:
            monitor_task = asyncio.create_task(self.monitor_resources())
            main_loop_task = asyncio.create_task(self.main_loop())

            await asyncio.gather(monitor_task, main_loop_task)
        except asyncio.CancelledError:
            Logger.log_info("Tasks have been cancelled due to shutdown.")
        finally:
            Logger.log_info("Script has been terminated.")

    def shutdown(self):
        Logger.log_info("Shutdown initiated. Setting shutdown flag.")
        self.shutdown_flag = True


def main():
    def signal_handler(sig, frame):
        print(f"Script interrupted by signal {sig}. Initiating graceful shutdown...")
        bot.shutdown()

    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")
    config_manager = ConfigManager(config_file=config_path)

    bot = RobloxAFKBot(config_manager)
    asyncio.run(bot.run())


if __name__ == "__main__":
    main()
