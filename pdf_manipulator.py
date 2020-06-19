from appJar import gui
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from pathlib import Path

# Initializing appJar GUI and defining window size
app = gui("PDF Manipulator", useTtk=True)
app.setTtkTheme()
app.setSize(500, 350)


def split_pages(input_file, page_range, out_file):
    # Page splitting function
    output = PdfFileWriter()
    input_pdf = PdfFileReader(open(input_file, "rb"))
    output_file = open(out_file, "wb")

    # https://stackoverflow.com/questions/5704931/parse-string-of-integer-sets-with-intervals-to-list
    page_ranges = (x.split("-") for x in page_range.split(","))
    range_list = [i for r in page_ranges for i in range(
        int(r[0]), int(r[-1]) + 1)]

    for p in range_list:
        # Need to subtract 1 because pages are 0 indexed
        try:
            output.addPage(input_pdf.getPage(p - 1))
        except IndexError:
            # Alert the user and stop adding pages
            app.infoBox(
                "Info", "Range exceeded number of pages in input.\nFile will still be saved.")
            break
    output.write(output_file)

    output_file.close()

    if app.questionBox("File Save", "Output PDF saved. Do you want to quit?"):
        app.stop()


def pdf_merge(first_file, second_file, out_file):
    # PDF merging function
    merger = PdfFileMerger()

    merger.append(first_file)
    merger.append(second_file)
    merger.write(out_file)

    # out_file.close()

    if app.questionBox("File Save", "Output PDF saved. Do you want to quit?"):
        app.stop()


def validate_split(input_file, output_dir, pg_range, file_name):
    # Validate that all required inputs have been provided for splitting
    errors = False
    error_msgs = []

    if Path(input_file).suffix.upper() != ".PDF":
        errors = True
        error_msgs.append("Please select an PDF input file!")

    if len(pg_range) < 1:
        errors = True
        error_msgs.append("Please enter a valid page range")

    if not (Path(output_dir)).exists():
        errors = True
        error_msgs.append("Please select a valid output directory")

    if len(file_name) < 1:
        errors = True
        error_msgs.append("Please enter a file name")

        return errors, error_msgs


def validate_merge(first_file, second_file, output_dir, file_name):
    # Validate that all required inputs have been provided for merging
    errors = False
    error_msgs = []

    if Path(first_file).suffix.upper() != ".PDF":
        errors = True
        error_msgs.append("Please select a PDF first input file!")

    if Path(second_file).suffix.upper() != ".PDF":
        errors = True
        error_msgs.append("Please select a PDF second input file!")

    if not (Path(output_dir)).exists():
        errors = True
        error_msgs.append("Please select a valid output directory")

    if len(file_name) < 1:
        errors = True
        error_msgs.append("Please enter a file name")

        return errors, error_msgs


def press_split(button):
    # Function for the splitting button
    if button == "Split!":
        src_file = app.getEntry("Split_Input_file")
        out_dir = app.getEntry("Split_Output_Directory")
        page_range = app.getEntry("Page_Ranges")
        out_file = app.getEntry("Split_Output_File")
        errors = validate_split(src_file, out_dir, page_range, out_file)
        error_msg = validate_split(src_file, out_dir, page_range, out_file)
        if errors:
            app.errorBox("Error", "\n".join(error_msg), parent=None)
        else:
            split_pages(src_file, page_range, str(
                Path(out_dir, out_file + ".pdf")))

    else:
        app.stop()


def press_merge(button):
    # Function for the merging button
    if button == "Merge!":
        src_file1 = app.getEntry("Merge_Input_file1")
        src_file2 = app.getEntry("Merge_Input_file2")
        out_dir = app.getEntry("Merge_Output_Dir")
        merged_out_file = app.getEntry("Merge_Output_name")
        errors = validate_merge(src_file1, src_file2, out_dir, merged_out_file)
        error_msg = validate_merge(
            src_file1, src_file2, out_dir, merged_out_file)
        if errors:
            app.errorBox("Error", "\n".join(error_msg), parent=None)
        else:
            pdf_merge(src_file1, src_file2, str(
                Path(out_dir, merged_out_file + ".pdf")))

    else:
        app.stop()


# Create a tabbed interface and define all elements
app.startTabbedFrame("Tabbed Frame")
app.startTab("Split")

app.addLabel("Choose Source PDF File")
app.addFileEntry("Split_Input_file")

app.addLabel("Select Output Directory")
app.addDirectoryEntry("Split_Output_Directory")

app.addLabel("Output file name")
app.addEntry("Split_Output_File")

app.addLabel("Page Ranges: 1,3,4-10")
app.addEntry("Page_Ranges")

app.addButtons(["Split!", "Quit"], press_split)
app.stopTab()

app.startTab("Merge")

app.addLabel("Choose First PDF File")
app.addFileEntry("Merge_Input_file1")

app.addLabel("Choose Second PDF File")
app.addFileEntry("Merge_Input_file2")

app.addLabel("Select Merged Output Directory")
app.addDirectoryEntry("Merge_Output_Dir")

app.addLabel("Merged Output File Name")
app.addEntry("Merge_Output_name")

app.addButtons(["Merge!", "Exit"], press_merge)

app.stopTab()

app.stopTabbedFrame()

app.go()
