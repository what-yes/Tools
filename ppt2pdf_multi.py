from genericpath import exists
import comtypes.client
import os
 
def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application", dynamic=True)
    powerpoint.Visible = 1
    return powerpoint
 
def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    print(outputFileName + "-------success")
    deck.Close()
 
def convert_files_in_folder(powerpoint, folder):
    files = os.listdir(folder)
    pdfdir = os.path.join(folder,"pdf")
    if not os.path.exists(pdfdir):
        os.mkdir(pdfdir)
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    for pptfile in pptfiles:
        fullpath = os.path.join(cwd, pptfile)
        filename = os.path.join(cwd,"pdf", pptfile.split(".")[0])
        ppt_to_pdf(powerpoint, fullpath, filename)
 
if __name__ == "__main__":
    powerpoint = init_powerpoint()
    # cwd = os.getcwd()
    cwd = "C:\大三上\网络信息安全\ppt"
    convert_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()