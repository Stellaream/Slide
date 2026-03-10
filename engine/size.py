import os
import win32com.client
import pythoncom 

def fix_ppt_with_drag_simulation(input_path):
    """
    通过 COM 接口模拟拖动，触发自动布局计算，并直接保存覆盖源文件。
    兼容 Flask 多线程环境，且执行后自动彻底关闭 PPT。
    """
    abs_input = os.path.abspath(input_path)

    if not os.path.exists(abs_input):
        print(f"❌ 未找到文件: {abs_input}")
        return

    print(f"正在调整字号并覆盖源文件: {os.path.basename(abs_input)} ...")
    
    # 初始化 COM (多线程必须)
    pythoncom.CoInitialize()
    ppt_app = None
    prs = None

    try:
        # 启动 PPT
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True 
        
        # 最小化窗口 (减少干扰)
        try:
            ppt_app.WindowState = 2 # ppWindowMinimized
        except:
            pass 
            
        # 打开文件
        try:
            prs = ppt_app.Presentations.Open(abs_input, WithWindow=True)
        except Exception as e:
            print(f"❌ 无法打开文件 (请确保文件未被占用): {e}")
            # 如果打开失败，尝试关闭应用，防止僵尸进程
            if ppt_app: ppt_app.Quit()
            return

        # 遍历并模拟拖动
        for slide in prs.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    tf = shape.TextFrame2
                    if tf.HasText:
                        # 模拟拖动触发重绘
                        shape.Left = shape.Left + 5 
                        shape.Left = shape.Left - 5
        
        # 保存
        prs.Save()
        print(f"调整完成！源文件已更新")

    except Exception as e:
        print(f"⚠️ 脚本运行出错: {e}")
        import traceback
        traceback.print_exc()
        
    finally:  
        # 1. 关闭演示文稿
        if prs:
            try:
                prs.Close()
            except:
                pass

        # 2. 退出 PowerPoint 软件
        if ppt_app:
            try:
                ppt_app.Quit()
                print("PowerPoint 已关闭")
            except:
                pass
        
        # 3. 释放 COM 环境
        pythoncom.CoUninitialize()