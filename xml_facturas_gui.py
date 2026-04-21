from __future__ import annotations

import csv
import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET


NS = {
    "inv": "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "ext": "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
    "sac": "urn:sunat:names:specification:ubl:peru:schema:xsd:SunatAggregateComponents-1",
}

HEADERS = [
    "# DE FACTURA",
    "PROVEEDOR",
    "CARROCERIA",
    "SERIE DE LA CARROCERIA",
    "MARCA",
    "MODELO",
    "AÑO MODELO",
    "PLACA",
    "# MOTOR",
    "# DE CHASIS",
    "# DE ASIENTOS",
    "PRECIO DE VENTA",
]


def text_or_empty(node: ET.Element | None) -> str:
    if node is None or node.text is None:
        return ""
    return node.text.strip()


def clean_prefix(raw_value: str, prefix: str) -> str:
    value = raw_value.strip()
    if value.upper().startswith(prefix.upper()):
        return value[len(prefix) :].strip()
    return value


def parse_description_fields(description: str) -> dict[str, str]:
    result = {
        "carroceria": "",
        "chasis": "",
        "motor": "",
        "anio_modelo": "",
    }

    if not description:
        return result

    first_line = description.splitlines()[0].strip()
    if first_line:
        result["carroceria"] = first_line.split(",")[0].strip()

    chasis_match = re.search(r"CHASIS\s*:?\s*:?\s*([A-Z0-9-]+)", description, re.IGNORECASE)
    if chasis_match:
        result["chasis"] = chasis_match.group(1).strip()

    motor_match = re.search(r"MOTOR\s*:?\s*:?\s*([A-Z0-9-]+)", description, re.IGNORECASE)
    if motor_match:
        result["motor"] = motor_match.group(1).strip()

    anio_match = re.search(r"A[ÑN]O\s+MODELO\s*:?\s*:?\s*(\d{4})", description, re.IGNORECASE)
    if anio_match:
        result["anio_modelo"] = anio_match.group(1).strip()

    return result


def parse_additional_properties(root: ET.Element) -> dict[str, str]:
    properties: dict[str, str] = {}
    for prop in root.findall(".//sac:AdditionalProperty", NS):
        prop_id = text_or_empty(prop.find("cbc:ID", NS))
        prop_value = text_or_empty(prop.find("cbc:Value", NS))
        if prop_id:
            properties[prop_id] = prop_value
    return properties


def parse_invoice(xml_path: Path) -> dict[str, str]:
    tree = ET.parse(xml_path)
    root = tree.getroot()

    additional = parse_additional_properties(root)

    invoice_number = text_or_empty(root.find("cbc:ID", NS))

    proveedor = text_or_empty(
        root.find("cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name", NS)
    )
    if not proveedor:
        proveedor = text_or_empty(
            root.find(
                "cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName",
                NS,
            )
        )

    description = text_or_empty(root.find("cac:InvoiceLine/cac:Item/cbc:Description", NS))
    description_fields = parse_description_fields(description)

    marca = clean_prefix(additional.get("02", ""), "MARCA:")
    modelo = clean_prefix(additional.get("04", ""), "MODELO:")
    motor = clean_prefix(additional.get("03", ""), "MOTOR:")
    chasis = clean_prefix(additional.get("07", ""), "NUMERO DE CHASIS:")

    if not chasis:
        chasis = description_fields["chasis"]
    if not motor:
        motor = description_fields["motor"]

    anio_modelo = description_fields["anio_modelo"]
    if not anio_modelo:
        anio_modelo = clean_prefix(additional.get("05", ""), "AÑO MODELO:")

    carroceria = description_fields["carroceria"] or "REMOLCADOR"

    total_igv = text_or_empty(root.find("cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount", NS))
    if not total_igv:
        total_igv = text_or_empty(root.find("cac:LegalMonetaryTotal/cbc:PayableAmount", NS))

    return {
        "# DE FACTURA": invoice_number,
        "PROVEEDOR": proveedor,
        "CARROCERIA": carroceria,
        "SERIE DE LA CARROCERIA": chasis,
        "MARCA": marca,
        "MODELO": modelo,
        "AÑO MODELO": anio_modelo,
        "PLACA": "",
        "# MOTOR": motor,
        "# DE CHASIS": chasis,
        "# DE ASIENTOS": "",
        "PRECIO DE VENTA": total_igv,
    }


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Lector XML Factura UBL")
        self.root.geometry("980x560")

        self.rows: list[dict[str, str]] = []

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")

        ttk.Label(
            top,
            text="Carga XML de factura y completa campos para Excel",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        buttons = ttk.Frame(top)
        buttons.pack(fill="x", pady=(10, 0))

        ttk.Button(buttons, text="Seleccionar XML", command=self.load_xml_files).pack(
            side="left"
        )
        ttk.Button(buttons, text="Exportar CSV", command=self.export_csv).pack(
            side="left", padx=8
        )
        ttk.Button(buttons, text="Limpiar", command=self.clear_rows).pack(side="left")

        table_frame = ttk.Frame(self.root, padding=(12, 8, 12, 12))
        table_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(table_frame, columns=HEADERS, show="headings")
        for col in HEADERS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=145, anchor="w")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def load_xml_files(self) -> None:
        xml_paths = filedialog.askopenfilenames(
            title="Selecciona uno o varios XML",
            filetypes=[("Archivos XML", "*.xml")],
        )
        if not xml_paths:
            return

        loaded = 0
        errors: list[str] = []

        for raw_path in xml_paths:
            path = Path(raw_path)
            try:
                data = parse_invoice(path)
                self.rows.append(data)
                self.tree.insert("", "end", values=[data[h] for h in HEADERS])
                loaded += 1
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")

        message = f"XML cargados: {loaded}"
        if errors:
            message += "\n\nErrores:\n" + "\n".join(errors[:8])
        messagebox.showinfo("Resultado", message)

    def export_csv(self) -> None:
        if not self.rows:
            messagebox.showwarning("Sin datos", "Primero carga al menos un XML.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Guardar CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="facturas_extraidas.csv",
        )
        if not save_path:
            return

        with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=HEADERS)
            writer.writeheader()
            writer.writerows(self.rows)

        messagebox.showinfo("Listo", f"CSV generado en:\n{save_path}")

    def clear_rows(self) -> None:
        self.rows.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)


def main() -> None:
    root = tk.Tk()
    app = App(root)
    _ = app
    root.mainloop()


if __name__ == "__main__":
    main()