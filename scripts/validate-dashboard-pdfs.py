from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

import fitz


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = PROJECT_ROOT / "app" / "data" / "diagnosticos.json"
PUBLIC_ROOT = PROJECT_ROOT / "public"
REPORT_PATH = PROJECT_ROOT / "outputs" / "diagnosticos-qa.json"


def main() -> int:
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    errors: list[dict[str, str]] = []
    checked: list[dict[str, object]] = []
    hashes: dict[str, str] = {}

    if len(manifest) != 158:
        errors.append(
            {
                "municipio": "inventario",
                "error": f"Se esperaban 158 documentos y hay {len(manifest)}.",
            }
        )

    for entry in manifest:
        municipality = entry["municipio"]
        pdf_path = PUBLIC_ROOT / entry["url"]
        try:
            size = pdf_path.stat().st_size
            if size < 500_000:
                raise ValueError(f"Archivo demasiado pequeño: {size} bytes")
            if size > 15_000_000:
                raise ValueError(f"Archivo demasiado grande: {size} bytes")

            digest = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
            if digest in hashes:
                raise ValueError(
                    f"PDF duplicado de {hashes[digest]} (mismo SHA-256)."
                )
            hashes[digest] = municipality

            document = fitz.open(pdf_path)
            page_count = document.page_count
            if not 7 <= page_count <= 10:
                raise ValueError(f"Número de páginas inesperado: {page_count}")

            page_texts = [page.get_text("text") for page in document]
            document.close()
            first_page = page_texts[0]
            all_text = "\n".join(page_texts)

            required = [
                municipality,
                "Diagnóstico Territorial",
                "Indicadores complementarios de población",
                "Resultados, riesgos e inversión territorial",
                "Resumen de Comparación",
            ]
            if entry.get("includesNarrative"):
                required.extend(["Resumen Narrativo", "Panorama general"])
            missing = [label for label in required if label not in all_text]
            if missing:
                raise ValueError(f"Texto requerido ausente: {', '.join(missing)}")
            if municipality not in first_page:
                raise ValueError("El nombre del municipio no aparece en la portada.")

            forbidden = [
                "Aún no se ha generado el resumen",
                "Error al generar resumen",
                "Cargando indicadores territoriales adicionales",
            ]
            found_forbidden = [label for label in forbidden if label in all_text]
            if found_forbidden:
                raise ValueError(
                    f"Texto de error presente: {', '.join(found_forbidden)}"
                )
            if entry.get("dashboardVersion") != "DDPT-Dashboard-Territorial-2026-08":
                raise ValueError("Versión del Dashboard ausente o incorrecta en el manifiesto.")

            checked.append(
                {
                    "id": entry["id"],
                    "municipio": municipality,
                    "adm2Code": entry["adm2Code"],
                    "filename": entry["filename"],
                    "bytes": size,
                    "pages": page_count,
                    "sha256": digest,
                    "dashboardVersion": entry.get("dashboardVersion"),
                }
            )
        except Exception as error:  # noqa: BLE001
            errors.append({"municipio": municipality, "error": str(error)})

    report = {
        "expected": 158,
        "checked": len(checked),
        "errors": errors,
        "documents": checked,
    }
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(
        f"PDF verificados: {len(checked)}/158. "
        f"Errores: {len(errors)}. Informe: {REPORT_PATH}"
    )
    if errors:
        for error in errors[:20]:
            print(f"- {error['municipio']}: {error['error']}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
