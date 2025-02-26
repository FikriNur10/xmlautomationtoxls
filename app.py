import xml.etree.ElementTree as ET
import os
import pandas as pd
from datetime import datetime

def extract_report_info(xml_path):
    try:
        with open(xml_path, "rb") as file:
            tree = ET.parse(file)

        root = tree.getroot()

        # Ambil nama file tanpa ekstensi dan format ulang jika diperlukan
        file_name = os.path.basename(xml_path).replace(".xml", "")
        # Mengubah formatted_name agar tidak ada tanda strip (-)
        formatted_name = file_name.split("--")[-1].replace("_", " ").replace("-", " ")

        summary_info = root.find("Summaryinfo")
        report_title = summary_info.get("ReportTitle", "").strip() if summary_info is not None else ""

        # Jika Report Name kosong, gunakan nama file
        if not report_title:
            report_title = formatted_name

        report_author = summary_info.get("ReportAuthor", "").strip() if summary_info is not None else ""
        last_update = datetime.now().strftime("%Y-%m-%d")

        database = root.find("Database/Tables/Table")
        data_source = database.get("Name", "").strip() if database is not None else ""

        connection_info = root.find("Database/Tables/Table/ConnectionInfo")
        database_name = connection_info.get("QE_DatabaseName", "").strip() if connection_info is not None else ""
        fields = [
            field.get("Name", "").strip()
            for field in database.findall("Fields/Field")
            if int(field.get("UseCount", "0")) > 0
        ] if database is not None else []
        formulas = [formula.get("FormulaName", "").strip() for formula in root.findall("DataDefinition/FormulaFieldDefinitions/FormulaFieldDefinition")]

        record_selection = root.find("DataDefinition/RecordSelectionFormula")
        filters = record_selection.text.strip() if record_selection is not None and record_selection.text else "None"

        groups = [group.get("ConditionField", "").strip() for group in root.findall("DataDefinition/Groups/Group")]
        summary = [summary.get("FormulaName", "").strip() for summary in root.findall("DataDefinition/SummaryFields/SummaryFieldDefinition")]
        parameters = [param.get("Name", "").strip() for param in root.findall("DataDefinition/ParameterFieldDefinitions/ParameterFieldDefinition")]

        subreports = root.find("SubReports")
        subreports = subreports.text.strip() if subreports is not None and subreports.text else "None"

        charts = root.find("ReportDefinition/Areas/Area[@Kind='ReportHeader']")
        charts = charts.text.strip() if charts is not None and charts.text else "None"

        return {
            "File Name": os.path.basename(xml_path),
            "Report Name": report_title,
            "Data Source": f"Database: {database_name}, Table: {data_source}",
            "Fields": ", ".join(fields) if fields else "None",
            "Filters": filters,
            "Formulas": ", ".join(formulas) if formulas else "None",
            "Parameters": ", ".join(parameters) if parameters else "None",
            "Groups": f"Group by {', '.join(groups)}" if groups else "None",
            "Subreports": subreports,
            "Charts": charts,
            "Summary": ", ".join(summary) if summary else "None",
            "Author": report_author,
            "Last Update": last_update,
            "Status": "Success"
        }

    except ET.ParseError:
        return {"File Name": os.path.basename(xml_path), "Status": "Failed - XML ParseError"}

    except Exception as e:
        return {"File Name": os.path.basename(xml_path), "Status": f"Failed - {str(e)}"}


def process_folder(folder_path):
    if not os.path.exists(folder_path):
        print(f"⚠️ Folder tidak ditemukan: {folder_path}")
        return

    xml_files = [f for f in os.listdir(folder_path) if f.endswith(".xml")]

    if not xml_files:
        print(f"⚠️ Tidak ada file XML di folder: {folder_path}")
        return

    output_excel = os.path.join(folder_path, "output.xlsx")
    failed_list = os.path.join(folder_path, "failed_list.txt")

    open(failed_list, "w").close()  # Kosongkan file gagal

    data_list = []
    success_count = 0
    failed_count = 0

    for file_name in xml_files:
        xml_path = os.path.join(folder_path, file_name)
        result = extract_report_info(xml_path)

        if "Success" in result["Status"]:
            success_count += 1
        else:
            failed_count += 1
            with open(failed_list, "a", encoding="utf-8") as file:
                file.write(f"{xml_path} - {result['Status']}\n")

        data_list.append(result)

    # Simpan hasil ke file Excel
    df = pd.DataFrame(data_list)
    df.to_excel(output_excel, index=False)

    print("\n===== Automation Test =====")
    print(f"✅ File berhasil diproses: {success_count}")
    print(f"❌ File gagal diproses: {failed_count}")
    print(f"📂 Output Excel: {output_excel}")
    print(f"📂 Daftar file gagal: {failed_list}")


# Contoh penggunaan
folder_path = "./rpt-xls2"  # Ganti dengan lokasi folder XML
process_folder(folder_path)
