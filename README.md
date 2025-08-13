# WorkTools

A few simple tools created for work for either: <br>
-Automation of Manual Tasks or <br>
-Comparison tools used to compare versions of files/data/reconciliation <br>

## Files

### RTF_FileExtract and 123Scan_Compare
-RTF_FileExtract > Extracts text from 123Scan.rtf file type and outputs text file inorder to compare text files using 123ScanCompare. <br>
-123Scan_Compare > Compares the two text files from RTF_FileExtract. Outputs a comparison html file which has side by side comparison highlighting differences.

### Compare_Final
-Reads in two excel files, compares values between the two files based on unique tag/key name (key_column), outputs only changes in file with the corresponding row number.

### Compare_WordFiles
-Reads and compares contends of two word files, compares line by line and writes differences to text file.

### DocumentID_Replace
-Replaces old document ids with new document ids based on mapping from an excel file.

### ExcelSpec_to_CSV
-Takes an excel file spec with multiple sheets and creates individual csv files, each file with the original sheet name and only containing the lines that are specified newly created or edited in the document. <br>
-Based on a specific use case with this particular spec document as only need the new/edited lines to be uploaded from the spec.

### PDF_to_Word
-Converts PDF file to word document for editing.

### PiVision_Export_Full
Logs into PI Dashboard URL, takes screenshot of the dashboard and saves output file, PowerAutomate then used to email the file to team members.

### Text_Compare
-Simple comparison of two text files and outputs html side by side comparison.

### XML_Compare
-Compares two xml files and outputs a browser link highlighting the differences side by side.
