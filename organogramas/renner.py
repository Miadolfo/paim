"""Organograma do time Renner (fonte: TIME_RENNER.pdf)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from organograma import BLACK, GRAY, ORANGE, RED, WINE, build  # noqa: E402

SUB = "O time de casa — Organograma time Renner"
ESCOPO = "Após o cargo: frentes de atuação (Renner BR, Renner Latam, LRSA, Conteúdo)"

slides = [
    {   # visão geral
        "title": "ORGANOGRAMA GERAL · TIME RENNER",
        "subtitle": SUB,
        "top": ("DIREÇÃO DE CONTA · CAROLINE LIMA", "Atendimento · BR · Latam · LRSA"),
        "columns": [
            {"title": "ATENDIMENTO", "color": GRAY, "items": [
                ("grid", 1, [("7 EXECUTIVAS DE CONTA", "BR · Latam · LRSA · Conteúdo", "summary"),
                             ("1 ASSISTENTE DE CONTA", "Conteúdo", "summary")])]},
            {"title": "PLANEJAMENTO", "color": RED, "items": [
                ("head", "EDUARDO FLORES", "VP de planejamento"),
                ("head", "PAOLA PASCOALE", "Head de planejamento"),
                ("grid", 1, [("+2 PESSOAS", "Executivo e assistente", "summary")])]},
            {"title": "MÍDIA", "color": WINE, "items": [
                ("head", "SIMONE MAGON", "Diretora de mídia"),
                ("head", "PALOMA KUSSLER", "Head de mídia"),
                ("grid", 1, [("+5 PESSOAS", "Planejamento de mídia", "summary")])]},
            {"title": "ANALYTICS E CRM", "color": BLACK, "items": [
                ("head", "LUCAS DINIZ", "Diretor de analytics e CRM"),
                ("head", "JOÃO MALOSSI", "Head de BI"),
                ("grid", 1, [("+7 PESSOAS", "BI, analytics e dados", "summary")])]},
            {"title": "CRIAÇÃO E CONTEÚDO", "color": ORANGE, "items": [
                ("head", "RODRIGO PINTO", "VP de criação"),
                ("head", "DEH BASTOS", "Diretora de criação e conteúdo"),
                ("grid", 1, [("+26 PESSOAS", "Criação, PDVX e conteúdo", "summary")])]},
        ],
        "legend": ["57 pessoas no total · detalhamento por área nos próximos slides", "line"],
    },
    {   # atendimento
        "title": "RENNER · ATENDIMENTO",
        "subtitle": SUB,
        "top": ("DIREÇÃO DE CONTA · CAROLINE LIMA", "Renner BR · Renner Latam · LRSA"),
        "columns": [
            {"title": "RENNER BR", "color": GRAY, "items": [
                ("grid", 1, [("JULIA ZENKER", "Executiva de conta · BR"),
                             ("RUTE SCHMITZ", "Executiva de conta · BR"),
                             ("AMANDA SANTOS", "Executiva de conta · BR"),
                             ("BRUNA FLORES", "Executiva de conta · BR · Latam")])]},
            {"title": "RENNER LATAM", "color": GRAY, "items": [
                ("grid", 1, [("DESIREE SEIXAS", "Executiva de conta · Latam")])]},
            {"title": "LRSA", "color": GRAY, "items": [
                ("grid", 1, [("CAROLINE TRILHA", "Executiva de conta · LRSA")])]},
            {"title": "CONTEÚDO", "color": GRAY, "items": [
                ("grid", 1, [("BRUNA BECKO", "Executiva de conta · Conteúdo BR · Latam"),
                             ("MARINA CASAGRANDE", "Assistente de conta · Conteúdo BR · Latam")])]},
        ],
        "legend": [ESCOPO, "line"],
    },
    {   # planejamento
        "title": "RENNER · PLANEJAMENTO",
        "subtitle": SUB,
        "top": ("VP DE PLANEJAMENTO · EDUARDO FLORES", "Renner BR · Renner Latam · LRSA"),
        "gray": [("PAOLA PASCOALE", "Head de planejamento · BR · Latam · LRSA")],
        "columns_width": 0.6,
        "columns": [
            {"title": "PLANEJAMENTO", "color": RED, "items": [
                ("grid", 2, [("ENIO XAVIER", "Executivo de planejamento · BR · Latam · LRSA"),
                             ("LUIZA MENEZES", "Assistente de planejamento · BR · Latam · LRSA")])]},
        ],
        "legend": [ESCOPO, "line"],
    },
    {   # mídia
        "title": "RENNER · MÍDIA",
        "subtitle": SUB,
        "top": ("DIREÇÃO DE MÍDIA · SIMONE MAGON", "Renner BR · Renner Latam · LRSA"),
        "gray": [("PALOMA KUSSLER", "Head de mídia · toda a frente digital")],
        "columns": [
            {"title": "PLANEJAMENTO DE MÍDIA", "color": WINE, "items": [
                ("grid", 3, [("LEANDRO FREITAS", "Planejador de mídia · BR"),
                             ("GABRIEL SOARES", "Planejador de mídia · BR"),
                             ("JULYE ALMEIDA", "Planejadora de mídia · BR"),
                             ("FERNANDA AZEVEDO", "Planejadora de mídia · BR · Latam"),
                             ("CARLA AZEVEDO", "Planejadora de mídia · toda a frente offline")])]},
        ],
        "legend": [ESCOPO, "line"],
    },
    {   # analytics e CRM
        "title": "RENNER · ANALYTICS E CRM",
        "subtitle": SUB,
        "top": ("DIREÇÃO DE ANALYTICS E CRM · LUCAS DINIZ", "Renner BR · Renner Latam · LRSA"),
        "gray": [("JOÃO MALOSSI", "Head de BI · BR · Latam")],
        "columns": [
            {"title": "BI E ANALYTICS", "color": BLACK, "items": [
                ("head", "EDUARDA FRIEDRICH", "Líder de BI · BR · Latam · LRSA"),
                ("grid", 3, [("THAMARA BASTOS OLIVEIRA", "Especialista de analytics · BR · Latam · LRSA"),
                             ("KEVI PEGORARO", "Cientista de dados · BR · Latam"),
                             ("ISABELA CANAL", "Analista de BI · BR · Latam"),
                             ("AMANDA OLIVEIRA", "Analista de BI · BR · Latam"),
                             ("JÚLIA DIAS", "Assistente de BI · BR · Latam"),
                             ("EDUARDO SCHOTKIS", "Assistente de BI · BR · Latam")])]},
        ],
        "legend": [ESCOPO, "line"],
    },
    {   # criação
        "title": "RENNER · CRIAÇÃO",
        "subtitle": SUB,
        "top": ("VP DE CRIAÇÃO · RODRIGO PINTO", "BR · Latam · LRSA"),
        "gray": [("DEH BASTOS", "Diretora de criação e conteúdo · BR · Latam · LRSA")],
        "columns": [
            {"title": "PDVX", "color": BLACK, "weight": 1, "items": [
                ("head", "ISMAEL GOLI", "Head de PDVX · BR · Latam"),
                ("grid", 1, [("BIANCA JANOWSKI", "Designer sênior · BR · Latam"),
                             ("JOSÉ PORTELA", "Designer sênior · BR · Latam")])]},
            {"title": "CRIAÇÃO", "color": ORANGE, "weight": 2, "items": [
                ("head", "RENATA MALINOSK", "Gerente de criação · BR · Latam · LRSA"),
                ("grid", 2, [("PEDRO CORTELETTI", "D.A. sênior · BR · Latam · LRSA"),
                             ("SAMUEL SALAMON", "Motion / D.A. pleno · BR · Latam · LRSA"),
                             ("IAGO FRANCISCO", "D.A. freela · BR · Latam · LRSA"),
                             ("GUILHERME LOURENÇO", "D.A. freela · BR · Latam · LRSA"),
                             ("LUANA VIEIRA SILVEIRA", "Redatora sênior · BR · Latam · LRSA"),
                             ("ANA LIVIA MARCILIO", "Redatora plena · BR · Latam · LRSA"),
                             ("ANE FRANKE", "Redatora plena · BR · Latam"),
                             ("VANESSA ARAUJO", "Redatora júnior/plena · BR · Latam · LRSA")])]},
        ],
        "legend": [ESCOPO, "line"],
    },
    {   # conteúdo
        "title": "RENNER · CONTEÚDO",
        "subtitle": SUB,
        "top": ("DIREÇÃO · DEH BASTOS", "Criação e conteúdo"),
        "gray": [("ANDRESSA MORAIS SILVA", "Supervisora de conteúdo · BR · Latam")],
        "columns": [
            {"title": "DIREÇÃO DE ARTE", "color": ORANGE, "weight": 2, "items": [
                ("grid", 2, [("FERNANDA KNAK", "D.A. sênior · Conteúdo BR · Latam"),
                             ("NICOLAS ROHLEDER", "D.A. pleno · Conteúdo BR · Latam"),
                             ("JOSI FLORES", "D.A. plena · Conteúdo BR · Latam"),
                             ("VINÍCIUS CAVALETTI", "D.A. pleno · Conteúdo BR · Latam"),
                             ("LUCCA BOL", "D.A. júnior · Conteúdo BR · Latam"),
                             ("MARIANNE ALVES SILVA", "D.A. júnior · Conteúdo BR · Latam"),
                             ("GABRIELA DE PAULO", "D.A. freela · Conteúdo BR · Latam")])]},
            {"title": "REDAÇÃO", "color": RED, "weight": 1, "items": [
                ("grid", 1, [("ALEXANDRA MENDEZ", "Redatora plena · Conteúdo Latam"),
                             ("JULIETA ALCANTARA", "Redatora plena · Conteúdo Latam"),
                             ("CAMILLA OLIVEIRA", "Redatora jr. e CM · Conteúdo BR · Latam"),
                             ("HANNAH JOB", "Redatora jr. e CM · Conteúdo BR · Latam")])]},
            {"title": "COMMUNITY", "color": BLACK, "weight": 1, "items": [
                ("grid", 1, [("LARISSA SANTOS DO Ó", "CM · Conteúdo BR · Latam"),
                             ("MARJORIE PAULA", "CM · Conteúdo BR · Latam")])]},
        ],
        "legend": [ESCOPO, "line"],
    },
]

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "Organograma_Time_Renner.pptx")
    print(build(slides, os.path.abspath(out), "Organograma Time Renner"))
