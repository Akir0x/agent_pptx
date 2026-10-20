# Agente 1: Catalogador

Transforma os slides da biblioteca em **moldes**: descreve cada slide, classifica
para que ele serve e dá nomes padronizados às caixas. O resultado alimenta o
agente 2, que monta as apresentações.

## Arquivos

| Arquivo | Para que serve |
| --- | --- |
| `extrair_slides.py` | Gera as imagens dos slides, a lista de caixas de cada um e a planilha vazia de levantamento |
| `instrucoes_agente.md` | Instruções para colar no agente do Copilot |
| `renomear_caixas.py` | Grava no PowerPoint os nomes novos das caixas a partir da planilha |

## Passo a passo

1. **Montar a biblioteca.** Copie para um único `biblioteca.pptx` os slides que
   vão virar moldes, todos com o mesmo tema. Deixe de fora prints de tela,
   gráficos colados como imagem e slides repetidos.

2. **Extrair.**
   ```
   python extrair_slides.py biblioteca.pptx
   ```
   Cria a pasta `saida/` com `miniaturas/`, `para_o_agente/`,
   `caixas_extraidas.xlsx` e `levantamento.xlsx` (vazia).

3. **Catalogar no Copilot.** No agente catalogador, mande de 3 a 5 slides por
   vez: as imagens de `miniaturas/` anexadas e o conteúdo dos `.txt` de
   `para_o_agente/` colado na mensagem.

4. **Colar na planilha.** Copie a tabela MOLDES da resposta para a aba
   `moldes` do `levantamento.xlsx` e a tabela CAIXAS para a aba `caixas`
   (sem repetir o cabeçalho).

5. **Revisar.** Procure as células com `REVISAR`, confira se nenhum
   `nome_novo` se repete no mesmo slide e se as descrições fazem sentido.

6. **Renomear as caixas.**
   ```
   python renomear_caixas.py biblioteca.pptx saida/levantamento.xlsx
   ```
   Gera `biblioteca_final.pptx`. O script avisa se alguma caixa não foi encontrada.

**Entrega para o agente 2:** `biblioteca_final.pptx` + `levantamento.xlsx`.

## Observações

- No Windows com PowerPoint instalado, as imagens saem pelo próprio PowerPoint.
  Sem ele, o script usa LibreOffice + `pdftoppm` (o visual pode variar um pouco).
- O limite de caracteres é uma estimativa pelo tamanho da caixa e da fonte.
  Vale testar os moldes mais usados digitando um texto longo.
- Rodar a extração de novo não apaga um `levantamento.xlsx` já preenchido.
