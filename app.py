import streamlit as st
from docxtpl import DocxTemplate
import tempfile
import os

# Cargo data dictionary
cargo_data = {
    "3, 5 - Dimethyl benzoyl chloride; DMBC": {
        "technicalName": "CORROSIVE LIQUID, ACIDIC, ORGANIC, N.O.S.. / (3, 5 – DIMETHYL BENZOYL CHLORIDE)",
        "class": "8",
        "unno": "3265",
        "subrisk": "8",
        "packingGroup": "II",
        "ems": "F-A, S-B",
        "flashPoint": "124 DEG. CEL",
        "marinePollutant": "YES",
        "qty": "1x20’ ISO TANK",
        "natureOfCargo": "(Solid,Liquid,Gas) - LIQUID",
        "emergencyTel": "1234567890",
        "emergencyContactPerson": "John Doe"
    }
}

# --- Load available templates from 'templates/' folder ---
def get_templates(folder="templates"):
    templates = {}
    if not os.path.exists(folder):
        st.error(f"Templates folder '{folder}' not found!")
        return templates
    for file in os.listdir(folder):
        if file.lower().endswith(".docx"):
            name = os.path.splitext(file)[0]
            templates[name] = os.path.join(folder, file)
    return templates

templates = get_templates()

st.title("Shipping Document Generator")

if not templates:
    st.warning("No Word templates found in 'templates/' folder. Please add .docx files.")
    st.stop()

selected_template_name = st.selectbox("Select Template", list(templates.keys()))
selected_template_path = templates[selected_template_name]

shippers = ['-', 'SHIVA PHARMACHEM LTD.PLOT NO. Z-88 & Z/88/4, SEZ PART-1,DAHEJ 392 130, TALUKA- VAGRA DISTRICT BHARUCH, GUJARAT, INDIA']
consignees = ['-', 'Corteva Agriscience International SàrlC/O Corteva Agriscience Italia SRLSS 11 KM 190.2,, , ,Mozzanica BG Italy, , 24050, Italy.Contact Person : Riva Giorgia']
ports = ['-', 'NHAVA SHEVA', 'GENOA', 'Port 3']
emergency_contacts = ['-', 'John Doe - 1234567890', 'Jane Smith - 9876543210']

cargo = st.selectbox("Select Cargo", ["-"] + list(cargo_data.keys()))

if cargo != "-" and cargo in cargo_data:
    details = cargo_data[cargo]
else:
    details = {k: "" for k in ["technicalName", "class", "unno", "subrisk", "packingGroup", "ems", "flashPoint",
                               "marinePollutant", "qty", "natureOfCargo", "emergencyTel", "emergencyContactPerson"]}

technical_name = st.text_input("Technical Name", details.get("technicalName", ""))
cargo_class = st.text_input("Class", details.get("class", ""))
unno = st.text_input("UNNO", details.get("unno", ""))
subrisk = st.text_input("Subrisk", details.get("subrisk", ""))
packing_group = st.text_input("Packing Group", details.get("packingGroup", ""))
ems = st.text_input("EMS", details.get("ems", ""))
flash_point = st.text_input("Flash Point", details.get("flashPoint", ""))
marine_pollutant = st.text_input("Marine Pollutant", details.get("marinePollutant", ""))
qty = st.text_input("QTY", details.get("qty", ""))
nature_of_cargo = st.text_input("Nature of Cargo", details.get("natureOfCargo", ""))
emergency_tel = st.text_input("24 hr Emergency Tel No. (at destination)", details.get("emergencyTel", ""))
emergency_contact_person = st.text_input("Emergency Contact Person (at destination)", details.get("emergencyContactPerson", ""))

shipper = st.selectbox("Shipper", shippers)
consignee = st.selectbox("Consignee", consignees)
pol = st.selectbox("Port of Loading (POL)", ports)
pod = st.selectbox("Port of Discharge (POD)", ports)
emergency_contact = st.selectbox("Emergency Contact Name/No", emergency_contacts)

outer_package = st.text_input("Outer Package *")
inner_package = st.text_input("Inner Package *")
gross_wt = st.text_input("Gross Weight *")
net_wt = st.text_input("Net Weight *")
equipment_type = st.text_input("Equipment Type *")

mandatory_fields_filled = all([
    outer_package.strip(),
    inner_package.strip(),
    gross_wt.strip(),
    net_wt.strip(),
    equipment_type.strip(),
])

if st.button("Generate Document"):
    if not mandatory_fields_filled:
        st.error("Please fill all mandatory fields marked with *")
    else:
        data = {
    "SHIPPER": shipper,
    "CONSIGNEE": consignee,
    "POL": pol,
    "POD": pod,
    "CARGO": cargo,
    "TECHNICAL_NAME": technical_name,
    "CLASS": cargo_class,
    "UNNO": unno,
    "SUBRISK": subrisk,
    "PACKING_GROUP": packing_group,
    "EMS": ems,
    "FLASH_POINT": flash_point,
    "MARINE_POLLUTANT": marine_pollutant,
    "QTY": qty,
    "NATURE_OF_CARGO": nature_of_cargo,
    "EMERGENCY_TEL": emergency_tel,
    "EMERGENCY_CONTACT_PERSON": emergency_contact_person,
    "EMERGENCY_CONTACT": emergency_contact,
    "OUTER_PACKAGE": outer_package,
    "INNER_PACKAGE": inner_package,
    "GROSS_WT": gross_wt,
    "NET_WT": net_wt,   
    "EQUIPMENT_TYPE": equipment_type,
}


        if not os.path.exists(selected_template_path):
            st.error(f"Template file '{selected_template_path}' not found!")
        else:
            doc = DocxTemplate(selected_template_path)
            doc.render(data)

            tmp_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            doc.save(tmp_docx.name)

            st.success("Document generated successfully!")
            st.download_button(
                "Download Filled Word Document",
                data=open(tmp_docx.name, "rb").read(),
                file_name=f"{selected_template_name}_filled.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
