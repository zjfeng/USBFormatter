import os
import ctypes
import win32api
import win32file
import win32con
import tkinter as tk
from tkinter import messagebox
import subprocess

def list_usb_drives():
    """列出所有连接的U盘"""
    drives = win32api.GetLogicalDriveStrings().split('\x00')[:-1]
    usb_drives = []
    for drive in drives:
        drive_type = win32file.GetDriveType(drive)
        if drive_type == win32file.DRIVE_REMOVABLE:
            try:
                volume_info = win32api.GetVolumeInformation(drive)
                usb_drives.append({
                    "drive": drive,
                    "label": volume_info[0],
                    "filesystem": volume_info[4],
                    "capacity": get_drive_capacity(drive)
                })
            except Exception as e:
                print(f"无法获取驱动器 {drive} 的信息: {e}")
    return usb_drives

def get_drive_capacity(drive):
    """获取驱动器容量"""
    sectors_per_cluster, bytes_per_sector, num_free_clusters, total_num_clusters = win32file.GetDiskFreeSpace(drive)
    total_capacity = (total_num_clusters * sectors_per_cluster * bytes_per_sector) / (1024 * 1024 * 1024)
    return f"{total_capacity:.2f} GB"

def format_drive(drive):
    """使用 subprocess 调用 Windows format 命令格式化指定驱动器"""
    try:
        # # 卸载驱动器
        # unmount_drive(drive)
        
        # 确保驱动器路径格式正确
        if not drive.endswith(":"):
            drive = drive + ":"

        print(f"正在格式化驱动器 {drive}...")
        # 调用 Windows format 命令
        command = f"echo.|format {drive} /FS:NTFS /Q /Y"
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            shell=True
        )
        print(f"格式化命令输出: {result.stdout}")
        print(f"格式化命令错误: {result.stderr}")
        if result.returncode != 0:
            raise Exception(result.stderr or result.stdout or "未知错误")
        print(f"驱动器 {drive} 格式化成功！")
        messagebox.showinfo("成功", f"驱动器 {drive} 格式化成功！")
    except Exception as e:
        print(f"格式化驱动器 {drive} 时出错: {e}")
        messagebox.showerror("错误", f"格式化驱动器 {drive} 时出错: {e}")

def format_drive_with_diskpart(drive):
    """使用 diskpart 工具格式化指定驱动器"""
    try:
        # 确保驱动器路径格式正确
        if not drive.endswith(":"):
            drive = drive + ":"

        print(f"正在格式化驱动器 {drive}...")
        # 创建 diskpart 脚本
        script = f"""
        select volume {drive[0]}
        format fs=ntfs quick
        exit
        """
        with open("diskpart_script.txt", "w") as f:
            f.write(script)

        # 调用 diskpart
        result = subprocess.run(
            ["diskpart", "/s", "diskpart_script.txt"],
            capture_output=True,
            text=True,
            shell=True
        )
        print(f"diskpart 输出: {result.stdout}")
        print(f"diskpart 错误: {result.stderr}")
        if result.returncode != 0:
            raise Exception(result.stderr or "未知错误")

        print(f"驱动器 {drive} 格式化成功！")
        messagebox.showinfo("成功", f"驱动器 {drive} 格式化成功！")
    except Exception as e:
        print(f"格式化驱动器 {drive} 时出错: {e}")
        messagebox.showerror("错误", f"格式化驱动器 {drive} 时出错: {e}")
    finally:
        # 删除临时脚本文件
        if os.path.exists("diskpart_script.txt"):
            os.remove("diskpart_script.txt")

def format_drive_with_pywin32(drive):
    """使用 pywin32 格式化指定驱动器"""
    try:
        # 确保驱动器路径格式正确
        if not drive.endswith(":"):
            drive = drive + ":"

        # 使用正确的路径格式
        drive_path = f"\\\\.\\{drive[0]}:"

        print(f"正在格式化驱动器 {drive_path}...")
        # 打开驱动器
        handle = win32file.CreateFile(
            drive_path,
            win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            0,
            None
        )

        # 卸载卷
        win32file.DeviceIoControl(
            handle,
            win32file.FSCTL_DISMOUNT_VOLUME,
            None,
            0
        )

        # 格式化卷（模拟调用 Windows 的格式化工具）
        win32api.FormatMessage(win32file.FSCTL_LOCK_VOLUME)

        print(f"驱动器 {drive} 格式化成功！")
        messagebox.showinfo("成功", f"驱动器 {drive} 格式化成功！")
    except Exception as e:
        print(f"格式化驱动器 {drive} 时出错: {e}")
        messagebox.showerror("错误", f"格式化驱动器 {drive} 时出错: {e}")

def unmount_drive(drive):
    """卸载驱动器"""
    try:
        if not drive.endswith(":"):
            drive = drive + ":"
        ctypes.windll.kernel32.SetVolumeMountPointW(drive, None)
        print(f"驱动器 {drive} 已卸载")
    except Exception as e:
        print(f"无法卸载驱动器 {drive}: {e}")

def main():
    global usb_drives  # 声明 usb_drives 为全局变量
    usb_drives = []  # 初始化为空列表

    root = tk.Tk()
    root.title("U盘格式化工具")
    root.geometry("500x400")

    def refresh_usb_list():
        """刷新 U 盘列表"""
        global usb_drives  # 使用全局变量
        usb_drives = list_usb_drives()  # 更新全局变量
        listbox.delete(0, tk.END)  # 清空列表框
        if not usb_drives:
            messagebox.showinfo("提示", "未检测到任何U盘！")
        else:
            for drive in usb_drives:
                listbox.insert(tk.END, f"驱动器: {drive['drive']}, 卷标: {drive['label']}, 文件系统: {drive['filesystem']}, 容量: {drive['capacity']}")

    # 创建滚动条和列表框
    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, font=("Arial", 10))
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar.config(command=listbox.yview)

    # 显示初始 U 盘信息
    refresh_usb_list()  # 初始化时加载 U 盘列表

    # 确认按钮
    def confirm_format():
        """确认并格式化所有列出的 U 盘"""
        # 检查列表是否为空
        if listbox.size() == 0:
            messagebox.showwarning("警告", "当前没有可用的 U 盘！")
            return

        # 弹出确认对话框
        confirm = messagebox.askyesno(
            "确认格式化",
            "确认格式化所有列出的 U 盘吗？"
        )
        if confirm:
            for drive in usb_drives:
                format_drive_with_diskpart(drive['drive'])
            messagebox.showinfo("完成", "所有 U 盘格式化完成！")
            refresh_usb_list()  # 格式化完成后刷新列表

    confirm_button = tk.Button(root, text="确认格式化", command=confirm_format, font=("Arial", 12))
    confirm_button.pack(pady=10)

    # 刷新按钮
    refresh_button = tk.Button(root, text="刷新", command=refresh_usb_list, font=("Arial", 12))
    refresh_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()