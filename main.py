import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from docx import Document
from lxml import etree
from tkinter import Tk, filedialog

# WordprocessingML namespace
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}
sys_reqs = ""

class SysReq:
    def __init__(self, req_id, req_desc):
        self.req_id = req_id
        self.req_desc = req_desc
        self.req_cover = []

    def print(self):
        print(f"System Requirement {self.req_id} - {self.req_desc}")
        for req in self.req_cover:
            req.print()

class Requirement:
    def __init__(self, module=None, iden=None, status=None, cat=None, safety=None,
                 ver_met=None, val_met=None, op=None, desc=None, cover=None):
        self.module = module
        self.iden = iden
        self.status = status
        self.cat = cat
        self.safety = safety
        self.ver_met = ver_met
        self.val_met = val_met
        self.op = op
        self.desc = desc
        self.cover = cover
        self.func_blocks = []

    def set_module(self, value):
        self.module = value

    def get_module(self):
        return self.module

    def set_iden(self, value):
        self.iden = value

    def get_iden(self):
        return self.iden

    def set_status(self, value):
        self.status = value

    def get_status(self):
        return self.status

    def set_cat(self, value):
        self.cat = value

    def get_cat(self):
        return self.cat

    def set_safety(self, value):
        self.safety = value

    def get_safety(self):
        return self.safety

    def set_ver_met(self, value):
        self.ver_met = value

    def get_ver_met(self):
        return self.ver_met

    def set_val_met(self, value):
        self.val_met = value

    def get_val_met(self):
        return self.val_met

    def set_op(self, value):
        self.op = value

    def get_op(self):
        return self.op

    def set_desc(self, value):
        self.desc = value

    def get_desc(self):
        return self.desc

    def append_decs(self, text):
        if self.desc:
            self.desc += "\n"
        self.desc += text

    def set_cover(self, value):
        self.cover = value

    def get_cover(self):
        return self.cover

    def print(self):
        print(f"\tIdentifier: {self.iden}" or f"\tidentifier: No identifier set")
        print(f"\t\tModule: {self.module}" or f"\tModule: No module set")
        print(f"\t\tIdentifier: {self.iden}" or f"\tidentifier: No identifier set")
        print(f"\t\tStatus: {self.status}" or f"\tStatus: No status set")
        print(f"\t\tCategory: {self.cat}" or f"\tCategory: No category set")
        print(f"\t\tSafety Level (DAL): {self.safety}" or f"\tSafety Level (DAL): No safety level set")
        print(f"\t\tVerification Method: {self.ver_met}" or f"\tVerification Method: No method set")
        print(f"\t\tValidation Method: {self.val_met}" or f"\tValidation Method: No method set")
        print(f"\t\tOperational Mode: {self.op}" or f"\tOperational Mode: No mode set")
        print(f"\t\tDescription: {self.desc}" or f"\tDescription: No description set")
        print(f"\t\tCoverage: {self.cover}" or f"\tCoverage: No coverage set")
        print(f"\t\tFunctional Blocks:")
        for f in self.func_blocks:
            print(f"\t\t\t- {f}")

def get_file(type):
    # Create a hidden root window
    root = Tk()
    root.withdraw()

    match type:
        case "System Requirements":
            title = "Select the file with System Requirements"
            filetypes = [
                ("Excel files", ".xlsx"),
                ("All files", ".*")
            ]
        case "FFRS":
            title = "Select the FFRS file"
            filetypes = [
                ("Word files", ".docx"),
                ("Word files with macro", ".docm"),
                ("All files", ".*")
            ]
        case "FAD":
            title = "Select the FAD file"
            filetypes = [
                ("Word files", ".docx"),
                ("Word files with macro", ".docm"),
                ("All files", ".*")
            ]
        case _:
            print("Selected type not recognized - Exit program")
            exit(1)

    # Open file dialog
    return filedialog.askopenfilename(title=title, filetypes=filetypes)

def load_reqs_from_sheet(filename):
    wb = load_workbook(filename)
    sheets = wb.sheetnames

    print("Available sheets:")
    for i, name in enumerate(sheets, start=1):
        print(f"{i}. {name}")

    while True:
        try:
            choice = int(input("\nEnter the number of the sheet you want to select: "))
            # choice = 4
            if 1 <= choice <= len(sheets):
                selected_sheet = sheets[choice - 1]
                break
            else:
                print("Please enter a valid number.")
        except ValueError:
            print("Please enter a number")

    return wb[selected_sheet]

def load_reqs_from_excel(filename, min_row=4, req_id_col=0, req_desc_col=1):
    reqs = []

    sheet = load_reqs_from_sheet(filename)

    # sys_req_spec_file = load_workbook(filename)
    # sys_req_spec_sheet = sys_req_spec_file[sheetname]

    for row in sheet.iter_rows(min_row=min_row, values_only=True):           # Skip header and empty rows
        req_id = row[req_id_col]         # Value in column A
        req_desc = row[req_desc_col]       # Value in column B

        if not req_id:
            continue

        if isinstance(req_desc, str):
            req_desc = req_desc.replace("\n", "")

        req = SysReq(req_id, req_desc)
        reqs.append(req)

    return reqs

def extract_text_from_docm_by_style(docm_path, target_styles):
    results = []                      # Only for debug purposes
    requirements = []
    currentReq = None

    with zipfile.ZipFile(docm_path, "r") as docx_zip:
        # Read the main document XML
        xml_content = docx_zip.read("word/document.xml")
        root = ET.fromstring(xml_content)

        # Iterate over all paragraphs <w:p>
        for p in root.findall(".//w:p", NS):
            # --- Check paragraph style ---
            p_style_el = p.find(".//w:pPr/w:pStyle", NS)
            paragraph_style = p_style_el.get("{%s}val" % NS["w"]) if p_style_el is not None else None

            # If paragraph style matches, extract all its text
            if paragraph_style in target_styles:
                text = "".join(t.text for t in p.findall(".//w:t", NS) if t.text)

                # To be use only for debug purposes
                # if text.strip():
                #    results.append((paragraph_style, text))

                match paragraph_style:
                    case 'ReqTag':
                        tag = [p.strip() for p in text.split("\uf0b7")]
                        currentReq = Requirement(
                            module=tag[0].strip("[]").split("-")[1].strip(" "),
                            iden=tag[0].strip("[]"),
                            status=tag[1],
                            cat=tag[2],
                            safety=tag[3],
                            ver_met=tag[4],
                            val_met=tag[5],
                            op=tag[6]
                        )
                    case 'ReqText':
                        if currentReq:
                            if currentReq.get_desc():
                                currentReq.append_decs(text)
                            else:
                                currentReq.set_desc(text)
                    case 'ReqCover':
                        if currentReq:
                            currentReq.set_cover(text.strip("[]").split(" ")[1])
                        requirements.append(currentReq)

    return results, requirements

def extract_tables_from_docm_xml(docm_path, target_headings=None, partial=False):
    """
    Extract tables directly from the XML of a .docm file and associate
    each table with the last heading that appears above it.

    Parameters
    ----------
    docm_path : str
        Path to the .docm file.
    target_headers : list or str or None
        Restrict to tables whose heading matches these strings.
    partial : bool
        If True, header match is substring-based.

    Returns
    -------
    List of dicts:
        [
          { "header": "<text>", "table": [[...], [...]] },
          ...
        ]
    """
    if isinstance(target_headings, str):
        target_headings = [target_headings]

    # --- 1) Read the XML from the .docm zip container ---
    with zipfile.ZipFile(docm_path) as z:
        xml_content = z.read("word/document.xml")

    # --- 2) Parse XML ---
    root = etree.fromstring(xml_content)
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    results = []
    current_header = None

    # --- 3) Iterate through child elements in order ---
    for child in root.xpath(".//w:body/*", namespaces=ns):

        # ---------------------------
        # A) Detect heading paragraphs
        # ---------------------------
        if child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p":

            # Paragraph text (concatenate all runs)
            texts = child.xpath(".//w:t/text()", namespaces=ns)
            para_text = "".join(texts).strip()

            # Style (Heading 1, Heading 2, etc.)
            pstyle = child.xpath(".//w:pStyle/@w:val", namespaces=ns)
            style_name = pstyle[0] if pstyle else ""

            if style_name.startswith("Titolo") and para_text:
                current_header = para_text

        # ---------------------------
        # B) Detect tables
        # ---------------------------
        elif child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl" and current_header:

            table_data = []
            for row in child.xpath(".//w:tr", namespaces=ns):
                cells = []
                for cell in row.xpath("./w:tc", namespaces=ns):
                    cell_text = "".join(cell.xpath(".//w:t/text()", namespaces=ns)).strip("[]")
                    cells.append(cell_text.strip("[]"))
                table_data.append(cells)

            if not table_data:
                continue

            first_row = " ".join(table_data[0]).lower()

            if partial:
                row_ok = any(f.lower() in first_row for f in target_headings)
            else:
                # exact cell match: any cell equals filter text
                row_ok = any(f.lower() == cell.lower() for f in target_headings for cell in table_data[0])

            if not row_ok:
                continue

            # Store result
            results.append({
                "header": current_header.strip(" Requirements traceability"),
                "table": table_data[1:]
            })

    return results

def check_coverage(reqs=None, tables=None):
    if tables and reqs:
        for t in tables:
            for r in reqs:
                if [r.iden] in t['table']:
                    r.func_blocks.append(t['header'])

    if reqs and sys_reqs:
        for r in reqs:
            for sr in sys_reqs:
                if sr.req_id == r.cover:
                    sr.req_cover.append(r.iden)

if __name__ == '__main__':
    # Select files to open
    SysReq_filename = get_file("System Requirements")
    FFRS_filename = get_file("FFRS")
    FAD_filename = get_file("FAD")

    # Parsing Excel file to extract data
    sys_reqs = load_reqs_from_excel(SysReq_filename)

    styles_to_extract = {
        "ReqTag",
        "ReqText",
        "ReqCover"
    }                               # List of styles to extract from the FFRS file
    texts, reqs = extract_text_from_docm_by_style(FFRS_filename, styles_to_extract)

    tables = extract_tables_from_docm_xml(FAD_filename, target_headings="Requirement Covered")

    check_coverage(reqs=reqs, tables=tables)
    for sr in sys_reqs:
        sr.print()
