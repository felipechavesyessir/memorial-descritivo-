# memorial-descritivo-

Automação para preencher um memorial descritivo `.docx` sem alterar a estrutura do arquivo, apenas substituindo os campos com `..........` pelos confrontantes extraídos de um PDF.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```bash
python preencher_memorial.py \
  --pdf /caminho/confrontantes.pdf \
  --template /caminho/memorial_template.docx \
  --output /caminho/memorial_preenchido.docx
```

O script tenta identificar confrontantes por rótulos como `NORTE`, `SUL`, `LESTE`/`NASCENTE`, `OESTE`/`POENTE` e substitui os placeholders na ordem:

1. Norte
2. Sul
3. Leste
4. Oeste
