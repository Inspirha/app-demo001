"""Regenera index.html leyendo los Excel de SUNARP en ./datos

Uso:  python build.py
Solo usa la biblioteca estándar de Python. Cuando SUNARP publique un Excel
nuevo, reemplaza el archivo en ./datos (mismo formato) y vuelve a ejecutar.
"""
import glob, json, os, re, zipfile, datetime
import xml.etree.ElementTree as ET

HERE = "."
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
FUENTES = {
    "transferencias": "transferencia-de-propiedad",
    "sid": "compra-venta-predial-sid",
}


def larga(p):
    """Prefijo para rutas largas en Windows."""
    p = os.path.abspath(p)
    return "\\\\?\\" + p if os.name == "nt" else p


def col(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + ord(ch) - 64
    return n - 1


def filas(path):
    z = zipfile.ZipFile(larga(path))
    ss = []
    for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
        ss.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    for r in ET.fromstring(z.read("xl/worksheets/sheet1.xml")).findall(".//m:row", NS):
        d = {}
        for c in r.findall("m:c", NS):
            v = c.find("m:v", NS)
            if v is not None:
                d[col(c.get("r"))] = ss[int(v.text)] if c.get("t") == "s" else float(v.text)
        yield d


def parsear(path):
    """Devuelve {departamento: {'YYYY-MM': valor}} (incluye TOTAL)."""
    out, meses = {}, None
    for d in filas(path):
        a = d.get(0)
        if not isinstance(a, str):
            continue
        if a.upper().startswith("DEPARTAMENTOS"):
            meses = [
                (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(d[i]))).strftime("%Y-%m")
                for i in range(1, 13) if isinstance(d.get(i), float)
            ]
        elif meses and not a.upper().startswith("NOTA") and any(isinstance(d.get(i), float) for i in range(1, 13)):
            for i, m in enumerate(meses):
                if isinstance(d.get(i + 1), float):
                    out.setdefault(a.strip().upper(), {})[m] = int(d[i + 1])
    return out


def main():
    datos, titulos = {}, {}
    for clave, patron in FUENTES.items():
        ruta = glob.glob(os.path.join(HERE, "datos", "*" + patron + "*.xlsx"))[0]
        bruto = parsear(ruta)
        # último mes con datos reales (los meses futuros vienen en 0)
        ultimo = max(m for m, v in bruto["TOTAL"].items() if v > 0)
        meses = sorted(m for m in bruto["TOTAL"] if m <= ultimo)
        deps = sorted(k for k in bruto if k != "TOTAL")
        datos[clave] = {
            "meses": meses,
            "total": [bruto["TOTAL"][m] for m in meses],
            "deps": {k: [bruto[k].get(m, 0) for m in meses] for k in deps},
        }
    with open(larga(os.path.join(HERE, "index.template.html")), encoding="utf-8") as f:
        html = f.read()
    html = html.replace("/*__DATA__*/null", json.dumps(datos, separators=(",", ":")))
    with open(larga(os.path.join(HERE, "index.html")), "w", encoding="utf-8") as f:
        f.write(html)
    for k, v in datos.items():
        print(k, v["meses"][0], "->", v["meses"][-1], len(v["deps"]), "departamentos")


if __name__ == "__main__":
    main()
