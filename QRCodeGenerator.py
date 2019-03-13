import os
import wx
import win32print
import win32ui
from PIL import Image, ImageWin

try:
    import qrcode
except ImportError:
    qrcode = None

try:
    import PyQRNative
except ImportError:
    PyQRNative = None

# -------------------------------------------------------------------------------------
class QRPanel(wx.Panel):
    def __init__(self, parent):
        """Constructor"""
        wx.Panel.__init__(self, parent=parent)
        self.photo_max_size = 240
        sp = wx.StandardPaths.Get()
        self.defaultLocation = ("C:\\Users\\dmcinerney\\Desktop\\Python Scripts\\QR Code Generator")


        # Document Type Label
        qrDataLbl1 = wx.StaticText(self, label="Document Type: ", size=(200, 20))
        self.qrDataTxt1 = wx.Choice(self, choices=['MembershipApplication',
                                                   'LoanApplication',
                                                   'ProofofID',
                                                   'ProofofAddress'], size=(200, 25))

        # Member Number Label
        qrDataLbl2 = wx.StaticText(self, label="Member Number: ", size=(200, 20))
        self.qrDataTxt2 = wx.TextCtrl(self, value="", size=(200, 25))

        # Document Date Label
        qrDataLbl3 = wx.StaticText(self, label="Document Date: ", size=(200, 20))
        self.qrDataTxt3 = wx.TextCtrl(self, value="", size=(200, 25))

        # Save Location
        defLbl = "Default save location: " + self.defaultLocation
        self.defaultLocationLbl = wx.StaticText(self, label=defLbl)

        # QR Button
        qrcodeBtn = wx.Button(self, label="Create QR Code")
        qrcodeBtn.Bind(wx.EVT_BUTTON, self.onUseQrcode)

        # Create Window Size
        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        qrDataSizer1 = wx.BoxSizer(wx.HORIZONTAL)
        qrDataSizer2 = wx.BoxSizer(wx.HORIZONTAL)
        qrDataSizer3 = wx.BoxSizer(wx.HORIZONTAL)
        qrBtnSizer = wx.BoxSizer(wx.VERTICAL)

        # Top Line
        self.mainSizer.Add(wx.StaticLine(self, wx.ID_ANY),
                           0, wx.ALL | wx.EXPAND, 5)

        # Labels & Boxes
        qrDataSizer1.Add(qrDataLbl1, 1, wx.ALL, 5)
        qrDataSizer1.Add(self.qrDataTxt1, 0, wx.ALL | wx.EXPAND, 5)
        self.mainSizer.Add(qrDataSizer1, 0, wx.ALL, 5)

        qrDataSizer2.Add(qrDataLbl2, 1, wx.ALL, 5)
        qrDataSizer2.Add(self.qrDataTxt2, 0, wx.ALL | wx.EXPAND, 5)
        self.mainSizer.Add(qrDataSizer2, 0, wx.ALL, 5)

        qrDataSizer3.Add(qrDataLbl3, 1, wx.ALL, 5)
        qrDataSizer3.Add(self.qrDataTxt3, 0, wx.ALL | wx.EXPAND, 5)
        self.mainSizer.Add(qrDataSizer3, 0, wx.ALL, 5)


        # Middle Line 1
        self.mainSizer.Add(wx.StaticLine(self, wx.ID_ANY),
                           0, wx.ALL | wx.EXPAND, 5)

        # Save Location
        self.mainSizer.Add(self.defaultLocationLbl, 0, wx.ALL, 5)

        # Middle Line 2
        self.mainSizer.Add(wx.StaticLine(self, wx.ID_ANY),
                           0, wx.ALL | wx.EXPAND, 5)

        # Create Button
        qrBtnSizer.Add(qrcodeBtn, 0, wx.ALL, 5)
        self.mainSizer.Add(qrBtnSizer, 0, wx.ALL | wx.CENTER, 0)

        self.SetSizer(self.mainSizer)
        self.Layout()

# ----------------------------------------------------------------------
    def onUseQrcode(self, event):
        """
        https://github.com/lincolnloop/python-qrcode
        """
        qr = qrcode.QRCode(version=1, box_size=5, border=4)
        qr.add_data(self.qrDataTxt1.GetStringSelection() + "|" + self.qrDataTxt2.GetValue() + "|" + self.qrDataTxt3.GetValue())
        qr.make(fit=True)
        x = qr.make_image()

        qr_file = os.path.join(self.defaultLocation, "C:\\Users\\dmcinerney\\Desktop\\Python Scripts\\Qr Code Generator\\QRCodes\\QRCode.bmp")
        img_file = open(qr_file, 'wb')
        x.save(img_file, 'BMP')
        img_file.close()


        # PRINT THE QR CODE
        HORZRES = 8
        VERTRES = 10

        # LOGPIXELS = dots per inch
        LOGPIXELSX = 88
        LOGPIXELSY = 90

        # PHYSICALWIDTH/HEIGHT = total area
        PHYSICALWIDTH = 110
        PHYSICALHEIGHT = 111

        # PHYSICALOFFSETX/Y = left / top margin
        PHYSICALOFFSETX = 112
        PHYSICALOFFSETY = 113

        printer_name = win32print.GetDefaultPrinter()
        file_name = "C:\\Users\\dmcinerney\\Desktop\\Python scripts\\QR Code Generator\\QRCodes\\QRCode.bmp"

        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
        printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
        printer_margins = hDC.GetDeviceCaps(PHYSICALOFFSETX), hDC.GetDeviceCaps(PHYSICALOFFSETY)

        bmp = Image.open (file_name)
        if bmp.size[0] > bmp.size[1]:
            bmp = bmp.rotate (90)

        ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
        scale = min (ratios)

        hDC.StartDoc(file_name)
        hDC.StartPage()

        dib = ImageWin.Dib(bmp)
        scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
        x1 = int((printer_size[0] - scaled_width) / 2)
        y1 = int((printer_size[1] - scaled_height) / 2)
        x2 = x1 + scaled_width
        y2 = y1 + scaled_height
        dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()

        # os.remove("C:\\Users\\dmcinerney\\Desktop\\Python Scripts\\QR Code Generator\\QRCodes\\QRCode.bmp")
# ----------------------------------------------------------------------

class QRFrame(wx.Frame):
    def __init__(self):
        """Constructor"""
        wx.Frame.__init__(self, None, title="QR Code Viewer", size=(450, 350))
        self.SetMinSize((450, 350))
        self.SetMaxSize((450, 350))
        panel = QRPanel(self)


if __name__ == "__main__":
    app = wx.App(False)
    frame = QRFrame()
    frame.Show()
    app.MainLoop()
