# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
files = ['Z02803', 'Z02804', 'Z05892', 'Z06018', 'Z02803', 'Z02804', 'Z05892', 'Z06018', 'Z06021', 'Z06022', 'Z06025', 'Z30778'
         , 'Z37055', 'Z37344', 'Z02478', 'Z07782', 'Z08400', 'Z03729', 'Z03730', 'Z02488', 'Z02530', 'Z05653', 'Z10639', 'Z10825',
         'Z10826', 'Z05090', 'Z06421', 'Z09129', 'Z14130', 'Z14133', 'Z01340', 'Z05893', 'Z17358', 'Z17359', 'Z09780', 'Z09792',
         'Z00896', 'Z01037', 'Z03162', 'Z04266']
import os
path = 'c:\\Users\\YuHsin\\Desktop\\凭证查询'
import tempfile
import win32api
import win32print
import win32ui
import zipfile
from PIL import Image, ImageWin


def printer_loading(filename):
    if(os.path.exists(filename)):
        print("Yes")
    open(filename, "r")
    win32api.ShellExecute(
        0,
        "print",
        filename,
        #
        # If this is None, the default printer will
        # be used anyway.
        #
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
    print(os.path)
    ct = 1
    if (os.path.exists(path)):
        print("exist!")
        for name in files:
            os.makedirs(path + "\\" + str(ct) + '_' + name)
            ct += 1;


def printer_pic(file_name):
    # 物理宽度、高度
    PHYSICALWIDTH = 110
    PHYSICALHEIGHT = 111
    # 物理偏移位置
    PHYSICALOFFSETX = 0
    PHYSICALOFFSETY = 0
    printer_name = win32print.GetDefaultPrinter()
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)
    printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
    # printer_margins = hDC.GetDeviceCaps (PHYSICALOFFSETX), hDC.GetDeviceCaps (PHYSICALOFFSETY)
    # 打开图片
    # ＃通过每个像素使它尽可能大
    # ＃页面不失真。
    bmp = Image.open(file_name)
    #bmp.show()
    if(bmp.size[0] > bmp.size[1]):
        #print('rotate...')
        #print(bmp.size)
        bmp = bmp.transpose(Image.ROTATE_90)
        #bmp.show()
        #print(bmp.size)
    ratios = [1.0 * 1754 / bmp.size[0], 1.0 * 1240 / bmp.size[1]]
    scale = min(ratios)
    # ＃开始打印作业，并将位图绘制到
    # ＃按比例缩放打印机设备。
    hDC.StartDoc(file_name)
    hDC.StartPage()
    dib = ImageWin.Dib(bmp)
    scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
    #print(scaled_width, scaled_height)
    x1 = int((printer_size[0] - scaled_width) / 2)
    y1 = int((printer_size[1] - scaled_height) / 2)
    # 横向位置坐标
    x1 = 0
    # 竖向位置坐标
    y1 = 0
    # 4倍为自适应图片实际尺寸打印
    x2 = x1 + bmp.size[0] * 4
    y2 = y1 + bmp.size[1] * 4
    dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    os.chdir('C:\\Users\\YuHsin\\Desktop\\凭证查询')
    ffs = os.listdir(os.getcwd())
    rr = 1
    for files in sorted(ffs):
        print(files)
        fs = os.listdir(os.getcwd()+ '\\' + files)
        for rt in sorted(fs):
            if(rt[-1] == 'g'):
                printer_pic(os.getcwd() + '\\' + files + '\\' + rt)
                print(rt)

    #printer_pic('C:\\Users\\YuHsin\\Desktop\\财务凭证\\Z06987\\1903010698700005.jpg')
    #print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
