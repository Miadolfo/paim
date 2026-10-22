"""Organogramas dos clientes da Marília: Camicado, Phyto Parfums, Kappesberg e Grupo K1 ESG."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from organograma import BLACK, GRAY, ORANGE, RED, WINE, build  # noqa: E402

GESTAO = ("GESTÃO DA CONTA · MINÉIA", "Gestão da conta")
DUDA = ("DUDA WOLF", "Atendimento sênior")
BRUNA = ("BRUNA CAVALHEIRO", "Atendimento júnior")
ISA = ("head", "ISA SILVEIRA", "Diretora de criação")
LEO = ("head", "LEO TORRES", "Líder de criação")


def area(title, color, people, weight=1):
    return {"title": title, "color": color, "weight": weight, "items": [("grid", 1, people)]}


slides = [
    {
        "title": "ORGANOGRAMA GERAL · TIME CAMICADO",
        "subtitle": "O time de casa — Organograma time Camicado",
        "top": GESTAO,
        "gray": [DUDA, BRUNA],
        "columns": [
            area("PLANEJAMENTO", RED, [("MATEUS TOURINHO", "Planejamento")]),
            area("MÍDIA", WINE, [("LUISA CANTELLI", "Mídia")]),
            {"title": "CRIAÇÃO", "color": ORANGE, "weight": 2, "items": [
                ISA,
                LEO,
                ("grid", 2, [("DANIEL", "Criação"),
                             ("JOÃO", "Criação"),
                             ("D.A.", "A definir", "hiring"),
                             ("REDAÇÃO", "A definir", "hiring")])]},
            area("PRODUÇÃO", BLACK, [("ANA PAULA LUCE", "Produção"), ("ALEX", "Produção")]),
            area("ESTÚDIO", GRAY, [("VANIS", "Estúdio")]),
        ],
        "legend": ["line", "hiring:A definir"],
    },
    {
        "title": "ORGANOGRAMA GERAL · TIME PHYTO PARFUMS",
        "subtitle": "O time de casa — Organograma time Phyto Parfums",
        "top": GESTAO,
        "gray": [BRUNA],
        "columns": [
            area("PLANEJAMENTO", RED, [("ROMULO", "Planejamento")]),
            area("MÍDIA", WINE, [("ERICA", "Mídia")]),
            {"title": "CRIAÇÃO", "color": ORANGE, "weight": 2, "items": [
                ISA,
                LEO,
                ("grid", 2, [("ADRIELI", "Criação"),
                             ("GABRIELA", "Criação"),
                             ("HELENA", "Criação"),
                             ("JOÃO", "Criação")])]},
            area("PRODUÇÃO", BLACK, [("ELISA", "Produção")]),
        ],
        "legend": ["line"],
    },
    {
        "title": "ORGANOGRAMA GERAL · TIME KAPPESBERG",
        "subtitle": "O time de casa — Kappesberg México e Argentina",
        "top": GESTAO,
        "gray": [DUDA],
        "columns_width": 0.85,
        "columns": [
            area("MÍDIA", WINE, [("RAUL", "Mídia")]),
            {"title": "CRIAÇÃO", "color": ORANGE, "weight": 2, "items": [
                ISA,
                LEO,
                ("grid", 2, [("RAQUEL LOSEKAN", "Redação"),
                             ("RAÍSSA", "Estagiária de D.A.")])]},
            {"title": "CRM E SEO", "color": BLACK, "items": [
                ("head", "VICTOR", "Gestão de CRM e SEO"),
                ("grid", 1, [("RACHEL", "CRM")])]},
        ],
        "legend": ["line"],
    },
    {
        "title": "ORGANOGRAMA GERAL · GRUPO K1 ESG",
        "subtitle": "O time de casa — Organograma Grupo K1 ESG",
        "top": GESTAO,
        "gray": [DUDA],
        "columns_width": 0.5,
        "columns": [
            {"title": "CRIAÇÃO", "color": ORANGE, "items": [
                ISA,
                ("grid", 2, [("D.A.", "Conforme disponibilidade", "hiring"),
                             ("REDAÇÃO", "Conforme disponibilidade", "hiring")])]},
        ],
        "legend": ["line", "hiring:Conforme disponibilidade"],
    },
]

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "Organogramas_Clientes_Marilia.pptx")
    print(build(slides, os.path.abspath(out), "Organogramas Clientes Marília"))
