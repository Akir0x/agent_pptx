# agent_pptx

Dois agentes do Copilot + scripts em Python para montar apresentações novas
reaproveitando o visual dos slides que a empresa já tem.

```
 slides existentes                                    apresentação nova
       │                                                     ▲
       ▼                                                     │
┌─────────────────────────┐   biblioteca_final.pptx   ┌──────────────────────────┐
│ AGENTE 1: CATALOGADOR   │   + levantamento.xlsx     │ AGENTE 2: MONTADOR       │
│ extrair_slides.py       │ ────────────────────────▶ │ agente no Copilot → JSON │
│ agente no Copilot       │                           │ montar_apresentacao.py   │
│ renomear_caixas.py      │                           │                          │
└─────────────────────────┘                           └──────────────────────────┘
   feito uma vez (e quando                               usado a cada nova
   entrarem slides novos)                                apresentação
```

| Pasta | O que tem |
| --- | --- |
| [`agente-1-catalogador/`](agente-1-catalogador/) | Transforma os slides em moldes: imagens, descrições, nomes das caixas e limites de texto |
| [`agente-2-montador/`](agente-2-montador/) | O agente final: escolhe os moldes, escreve o conteúdo e gera o `.pptx` |
| [`exemplo/`](exemplo/) | Biblioteca de teste com 5 moldes e o levantamento já preenchido |

Não usa API de IA: a parte inteligente roda em agentes do Copilot, e os
scripts fazem a parte mecânica (ler e escrever o PowerPoint).

## Instalação

Python 3.10 ou mais novo.

```
pip install -r requirements.txt
```

No Windows com PowerPoint instalado, as imagens dos slides saem pelo próprio
PowerPoint. Em outro sistema, instale o LibreOffice e o `pdftoppm` (poppler).

## Testar o fluxo completo com o exemplo

```
cd exemplo
python criar_biblioteca_exemplo.py
python ../agente-1-catalogador/extrair_slides.py biblioteca.pptx
python ../agente-1-catalogador/renomear_caixas.py biblioteca.pptx levantamento_exemplo.xlsx
python ../agente-2-montador/montar_apresentacao.py ../agente-2-montador/exemplo/pedido_exemplo.json --biblioteca biblioteca_final.pptx --levantamento levantamento_exemplo.xlsx
```

O último comando gera `resultados_q3_cliente_exemplo.pptx` com 6 slides.
O `levantamento_exemplo.xlsx` faz o papel do que o agente 1 devolveria.
