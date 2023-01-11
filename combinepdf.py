print("""██████╗ ██████╗ ███████╗     ██████╗ ██████╗ ███╗   ██╗██╗   ██╗███████╗██████╗ ███████╗██╗ ██████╗ ███╗   ██╗     █████╗ ███╗   ██╗██████╗     ███╗   ███╗███████╗██████╗  ██████╗ ██╗███╗   ██╗ ██████╗ 
██╔══██╗██╔══██╗██╔════╝    ██╔════╝██╔═══██╗████╗  ██║██║   ██║██╔════╝██╔══██╗██╔════╝██║██╔═══██╗████╗  ██║    ██╔══██╗████╗  ██║██╔══██╗    ████╗ ████║██╔════╝██╔══██╗██╔════╝ ██║████╗  ██║██╔════╝ 
██████╔╝██║  ██║█████╗      ██║     ██║   ██║██╔██╗ ██║██║   ██║█████╗  ██████╔╝███████╗██║██║   ██║██╔██╗ ██║    ███████║██╔██╗ ██║██║  ██║    ██╔████╔██║█████╗  ██████╔╝██║  ███╗██║██╔██╗ ██║██║  ███╗
██╔═══╝ ██║  ██║██╔══╝      ██║     ██║   ██║██║╚██╗██║╚██╗ ██╔╝██╔══╝  ██╔══██╗╚════██║██║██║   ██║██║╚██╗██║    ██╔══██║██║╚██╗██║██║  ██║    ██║╚██╔╝██║██╔══╝  ██╔══██╗██║   ██║██║██║╚██╗██║██║   ██║
██║     ██████╔╝██║         ╚██████╗╚██████╔╝██║ ╚████║ ╚████╔╝ ███████╗██║  ██║███████║██║╚██████╔╝██║ ╚████║    ██║  ██║██║ ╚████║██████╔╝    ██║ ╚═╝ ██║███████╗██║  ██║╚██████╔╝██║██║ ╚████║╚██████╔╝
╚═╝     ╚═════╝ ╚═╝          ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝  ╚═══╝  ╚══════╝╚═╝  ╚═╝╚══════╝╚═╝ ╚═════╝ ╚═╝  ╚═══╝    ╚═╝  ╚═╝╚═╝  ╚═══╝╚═════╝     ╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝ """)


print("v1.0")
print("")
print("")
print("By-\n Tanmay Sheoran")

print("")
print("")
try:
    import os
    import comtypes.client
    from PyPDF2 import PdfMerger
    from tqdm import tqdm

    print("Currently supported file extensions: '.pdf', '.pptx', '.pptm', '.docx', '.xlsx', '.xlsm'")
    print("------------------------------------------------------------------------------------------------------------------------------------------------------")
    directory = input("Enter directory for files to convert and merge: ")
    directory = r"{0}".format(directory)
    directory2 = f'{directory}\\tempfiles'
    if not os.path.exists(directory2):
        os.makedirs(directory2)
    final_output_filename = input("Enter output filename: ")
    if final_output_filename[-3:] != 'pdf':
        final_output_filename = final_output_filename + ".pdf"


    def spinning_cursor():
        while True:
            for cursor in '|/-\\':
                yield cursor


    def ppt_2_pdf(input_ppt_file, output_pdf_file, format_type=32):
        """
        Convert a Powerpoint file to a pdf file
        :param input_ppt_file: input Powerpoint file
        :param output_pdf_file: output pdf file
        :param format_type:
        :return: a pdf file written in the directory
        """
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1

        if output_pdf_file[-3:] != 'pdf':
            output_pdf_file = output_pdf_file + ".pdf"

        ppt_file = powerpoint.Presentations.Open(input_ppt_file)
        ppt_file.SaveAs(output_pdf_file, format_type)
        ppt_file.Close()
        powerpoint.Quit()

    def convert_all_ppt(directory):
        """
        Convert all Powerpoint files in the same directory to pdf files
        :param directory: the full path to the directory
        :return: all pdf outputs in the same directory
        """
        try:
            for file in os.listdir(directory):
                _, file_extension = os.path.splitext(file)
                if "ppt" in file_extension:
                    input_file = directory + "\\" + file
                    output_file = input_file + "_output.pdf"
                    ppt_2_pdf(input_file, output_file)
        except FileNotFoundError:
            print("The system cannot file directory \'{0}\'".format(directory))
            exit(2)
        except OSError:
            print("The filename, directory name, or volume label syntax is incorrect: \'{0}\'"
                .format(directory))
            exit(2)

    file_path_list = [] 
    file_list = []
    # iterate over files in
    # that directory
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        # checking if it is a file
        if os.path.isfile(f):
            file_list.append(filename)
            file_path_list.append(f)

    print(f"Total no. of files to merge: {len(file_list)}")

    print("------------------------------------------------------------------------------------------------------------------------------------------------------")
    print("")
    count = 0
    ppts = []
    otherpdfs = []
    wdFormatPDF = 17
    files_not_converted = []
    items_to_convert = 0

    for i in file_path_list:
        if not i.endswith("pdf"):
            otherpdfs.append(i)
            items_to_convert+=1

    print(f"Non PDF files to convert: {items_to_convert}")

    for i in tqdm(otherpdfs):
        try:
            filename = directory2 +f"\\{count}.pdf"
            if(i.endswith("pptx") or i.endswith("pptm")):
                ppt_2_pdf(i,filename)
                count+=1  
            
            if(i.endswith(".docx")):
                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(i)
                doc.SaveAs(filename, FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()
                count+=1

            if(i.endswith(".xlsx") or i.endswith(".xlsm")):
                excel = comtypes.client.CreateObject("Excel.Application")
                sheets = excel.Workbooks.Open(i)
                for worksheet in sheets.Worksheets:
                    try:
                        worksheet.ExportAsFixedFormat(0, filename)
                        count+=1
                    except:
                        continue    
                excel.Close(True)    
        except Exception as e:
            files_not_converted.append(i)       

    for j in otherpdfs:
        file_path_list.remove(j)

    print("Conversion Completed.")

    print("")

    print("------------------------------------------------------------------------------------------------------------------------------------------------------")
    print(f"Merging PDFs together to {final_output_filename} ...")
    print("")
    merger = PdfMerger()

    for filename in os.listdir(directory2):
        f = os.path.join(directory2, filename)
        # checking if it is a file
        file_path_list.append(f)

    for pdf in tqdm(file_path_list):
        try:
            merger.append(pdf)
        except:
            files_not_converted.append(pdf)  
    print("Saving File...")
    merger.write(final_output_filename)
    merger.close()

    print(f"Files merged and save to ouput: {final_output_filename}")
    print("")
    print("------------------------------------------------------------------------------------------------------------------------------------------------------")
    print("")
    if(len(files_not_converted) > 0):
        print("Files not converted/merged due to error: ")
        for file in files_not_converted:
            print(file)
    print("")
    print("------------------------------------------------------------------------------------------------------------------------------------------------------")
    print("Thank you.")        
    input("Press Enter To Close...")        
except Exception as e:
    print(e)
    input("Press Enter to continue...")          