from genericpath import exists
import comtypes.client
import os
 
def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application", dynamic=True)
    powerpoint.Visible = 1
    return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    print(outputFileName + "-------success")
    deck.Close()

if __name__ == "__main__":
    powerpoint = init_powerpoint()
    # cwd = os.getcwd()
    pptfile = "C:\大三上\网络信息安全\ppt\crypto-0.ppt"
    pdffile = pptfile.split(".")[0] + ".pdf"
    ppt_to_pdf(powerpoint, pptfile, pdffile)
    powerpoint.Quit()