# Montador 100% no Copilot (interpretador de código)

Nesta versão o próprio agente gera o `.pptx`, sem ninguém rodar Python no
computador. Ele usa o **interpretador de código** do Copilot Studio, que roda o
código de `montador_compacto.py` (o mesmo método do `montar_apresentacao.py`,
resumido para caber nas instruções).

> Custo: o interpretador de código é um recurso premium e consome
> **créditos do Copilot**. Mensagens no chat de teste do Copilot Studio não são cobradas.

## Arquivos

| Arquivo | Para que serve |
| --- | --- |
| `instrucoes_agente.txt` | Cole inteiro no campo **Instruções** do agente (6.262 caracteres; o limite é 8.000) |
| `montador_compacto.py` | O código que vai embutido nas instruções, separado para manutenção e teste |

Se mudar o `montador_compacto.py`, gere o `instrucoes_agente.txt` de novo
colando o código no fim, depois de `CÓDIGO`.

## Configuração no Copilot Studio

1. Crie o agente e cole o `instrucoes_agente.txt` em **Instruções**.
2. Em **Conhecimento**, adicione o `levantamento.xlsx` preenchido.
3. Em **Configurações > IA generativa**, ligue o **interpretador de código**
   e o upload de arquivos.
4. Em **Segurança > Autenticação**, use autenticação de usuário (o interpretador
   não funciona com "sem autenticação").

## Como o usuário usa

1. Pede a apresentação e manda o conteúdo.
2. O agente mostra o roteiro; o usuário ajusta se quiser.
3. O usuário **anexa o `biblioteca_final.pptx`** e escreve "gerar".
4. O agente devolve o `apresentacao.pptx` pronto para baixar.

O anexo vai só na última mensagem porque o interpretador não mantém um arquivo
anexado entre várias mensagens.

## O que testar primeiro

Esta parte não foi testada dentro do Copilot, só localmente. Confira, nesta ordem:

1. O agente aceita o `.pptx` como anexo e o código consegue abrir o arquivo.
2. O `python-pptx` está disponível no interpretador (no teste publicado pela
   Classmethod estava).
3. O código roda sem ser alterado pelo agente. Se ele "reescrever" o código,
   reforce a regra nas instruções.
4. O arquivo devolvido abre no PowerPoint com o design dos moldes.

Se o passo 1 falhar, a alternativa é o agente chamar um fluxo do Power Automate
que roda o `montar_apresentacao.py` (Power Automate Desktop), sem gastar créditos.
