import subprocess
import os

# Define file paths (update these to match your actual files)
# input_docx: str = r"56.1062_S00R00_Zone1_PCN_SCL_20250218.docx"
input_docx: str = r"56.1062_Well14_PCN_20250131.docx"
vbs_script: str = r"convert_word_to_pdf.vbs"

input_docx = os.path.normpath(os.path.abspath(input_docx))
output_pdf: str = f"{input_docx.removesuffix(r".docx")}" + r".pdf"
output_pdf = os.path.normpath(os.path.abspath(output_pdf))
vbs_script = os.path.normpath(os.path.abspath(vbs_script))

# Run the VBS script with input/output arguments
subprocess.run(["cscript", vbs_script, input_docx, output_pdf], shell=True)
