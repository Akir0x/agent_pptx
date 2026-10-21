# Agente 2: Montador (agente final)

Cria apresentações novas reaproveitando os moldes catalogados pelo agente 1.

Há duas formas de usar:

| Forma | Quem gera o `.pptx` | Custo |
| --- | --- | --- |
| [`copilot-interpretador/`](copilot-interpretador/) | O próprio agente, com o interpretador de código do Copilot Studio | Consome créditos do Copilot |
| Esta pasta (`instrucoes_agente.md` + `montar_apresentacao.py`) | O agente devolve um JSON e o script em Python monta o arquivo | Sem créditos de IA |

O restante deste README descreve a forma com o script em Python.

## Arquivos

| Arquivo | Para que serve |
| --- | --- |
| `instrucoes_agente.md` | Instruções para colar no agente do Copilot |
| `montar_apresentacao.py` | Lê o JSON do agente e gera o PowerPoint |
| `exemplo/pedido_exemplo.json` | Exemplo de pedido no formato que o agente devolve |

## O que precisa estar pronto

Do agente 1: `biblioteca_final.pptx` (com as caixas renomeadas) e
`levantamento.xlsx` preenchido. O levantamento também vai como fonte de
conhecimento do agente no Copilot.

## Passo a passo

1. Converse com o agente montador: diga o objetivo, o público e mande o
   conteúdo (textos, números, documentos).
2. Revise o roteiro que ele propõe e responda "pode gerar".
3. Copie o bloco JSON da resposta e salve como `pedido.json`. Pode colar com a
   cerca ```` ```json ```` mesmo, o script ignora.
4. Rode:
   ```
   python montar_apresentacao.py pedido.json --biblioteca biblioteca_final.pptx --levantamento levantamento.xlsx
   ```
5. Abra o arquivo gerado e veja os avisos que o script mostrou no terminal.

## O que o script faz

- Copia cada molde pedido (mantendo fundo, cores, imagens e gráficos) e apaga
  os moldes que não foram usados.
- Preenche as caixas pelo nome, mantendo fonte, cor e tamanho do molde.
- Troca os dados de gráficos nativos e o texto de tabelas.
- Troca imagens quando o campo traz um caminho de arquivo que existe; quando
  traz uma sugestão (`[sugestão: ...]`), mantém a imagem e anota nas notas.
- Avisa quando um texto passa do limite, quando uma caixa não existe no molde
  e quando sobrou caixa sem conteúdo.
