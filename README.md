# Suíte de Testes — InfrasFileManager

## Estrutura

```
tests/
  test_emission.py         # 50 testes unitários — classe Emission
  test_excel_functions.py  # 16 testes unitários — excel_functions.py
  test_integration.py      #  6 testes de integração — filesystem + Excel real
  fixtures/
    IFS-XXXX-XXX-X-LD-XXXX.xlsx   # planilha LD (template para testes de integração, sem macros)
    IFS-XXXX-XXX-X-LD-XXXX.xlsm   # planilha LD (template com macros)
```

## Como rodar

```bash
# Todos os testes unitários (sem Excel):
python -m unittest discover tests

# Apenas integração (requer Windows + Excel instalado):
python -m unittest tests.test_integration -v
```

---

## `test_emission.py` — Classe `Emission`

Usa `object.__new__(Emission)` para instanciar sem `__init__`, evitando GUI, Excel e filesystem.

### `TestVerifyPattern`
- nome válido sem revisão, disciplina 1 letra: `IFS-2227-001-A-AR-00001.pdf` → `True`
- nome válido com revisão: `IFS-2227-001-A-AR-00001_R1.pdf` → `True`
- nome válido com disciplina 3 letras: `IFS-2227-001-GER-LD-00001.pdf` → `True`
- disciplina com 2 letras: `IFS-2227-001-GE-LD-00001.pdf` → `False`
- sem prefixo IFS: `documento.pdf` → `False`
- número de projeto curto: `IFS-22-001-A-AR-00001.pdf` → `False`

### `TestVerifyDatePattern`
- `"01/04/26"` → `True`
- `"1/4/26"` → `False`
- `""` → `False`
- `None` → levanta `TypeError`

### `TestVerifyLdPatternNoRev`
- LD válida sem revisão: `IFS-2227-001-GER-LD-00001` → `True`
- LD válida com `_R0`: `IFS-2227-001-GER-LD-00001_R0` → `True`
- disciplina com 2 letras → `False`
- nome genérico → `False`

### `TestGetFolderName`
- `IFS-2227-001-A-AR-00001_R2.pdf` → `IFS-2227-001-A-AR-00001` (23 chars)
- `IFS-2227-001-A-AR-00001.pdf` (sem revisão) → `IFS-2227-001-A-AR-00001` (23 chars)
- `IFS-2227-001-GER-LD-00001_R0.xlsx` (disciplina 3 letras) → `IFS-2227-001-GER-LD-00001` (25 chars)
- `IFS-2227-001-GER-LD-00001.xlsx` (sem revisão, disciplina 3 letras) → `IFS-2227-001-GER-LD-00001` (25 chars)

### `TestGetRevision`
- `IFS-2227-001-A-AR-00001_R2.pdf` → `2`
- sem revisão → `0`
- `_R0` → `0`
- `_R15` → `15`

### `TestGetFileName`
- `IFS-2227-001-A-AR-00001_R2.pdf` → `IFS-2227-001-A-AR-00001`
- sem revisão → nome sem extensão
- `IFS-2227-001-GER-LD-00001_R0.xlsx` (disciplina 3 letras) → `IFS-2227-001-GER-LD-00001`

### `TestGetLdRevision`
- `IFS-2227-001-GER-LD-00001_R3.xlsx` → `3`
- nome inválido → `-1`
- placeholder `IFS-XXXX-XXX-X-LD-XXXX.xlsx` → `-1`

### `TestGetRegExpressions`
- env vars definidas → retorna valores do env
- sem env vars → retorna defaults hardcoded

### `TestGetFileNumCaract`
- `FILE_NUM_CARACT=30` → `30` (int)
- sem env var → `23`

### `TestGetEmitedPath`
- estrutura com `3_Emitidos` → retorna path correto
- sem `3_Emitidos` → levanta `FileNotFoundError`

### `TestGetLdPath`
- `_LDs` dentro de `emited_path` → retorna esse path
- sem `_LDs` mas com `00_LDs` no cwd → retorna `00_LDs`
- nenhum dos dois → `FileNotFoundError`

### `TestGetFiles`
- 2 arquivos únicos → lista com 2 dicts corretos (`file_name`, `rev`, `emit`, `subdir`)
- arquivo duplicado → chama `msgbox` e `sys.exit`
- arquivos ocultos (`.arquivo`) → ignorados

### `TestCheckFilenamePattern`
- docs com nomes válidos → todos `emit=True`, `text_box` não chamado
- nome inválido → `emit=False` e `text_box` chamado uma vez

### `TestIssuedDirectories`
- doc cujo folder não existe → entra em `dirs_to_create`
- doc cujo folder já existe → chama `check_file`

### `TestConfirmFiles`
- usuário desmarca 1 arquivo → `emit=False` para o desmarcado
- 1 único arquivo → chama `ccbox`
- lista vazia → `msgbox` e `sys.exit`

### `TestDuplicatedFile`
- "Não emitir esse arquivo" → `doc['emit'] = False`
- "Emitir mesmo assim" → cria pasta `Obsoleto`, move arquivo, `doc['emit'] = True`
- "Cancelar" → `sys.exit`

---

## `test_excel_functions.py` — `excel_functions.py`

### `TestGetCoverCell`

| Revisão | Resultado  |
|---------|------------|
| 0       | `[32, 3]`  |
| 1       | `[32, 5]`  |
| 2       | `[32, 7]`  |
| 3       | `[32, 8]`  |
| 4       | `[32, 11]` |
| 5       | `[37, 3]`  |
| 6       | `[37, 5]`  |
| 7       | `[37, 7]`  |
| 8       | `[37, 8]`  |
| 9       | `[37, 11]` |
| 15      | `[37, 11]` |

### `TestGetGrdNumber`
- planilha com `GRD-001` e `GRD-002` (além da aba template `GRD-XXX`) → retorna `3`
- planilha com apenas a aba template `GRD-XXX` → retorna `1`

### `TestGetAcronymDefaultList`
- lê células das linhas imediatamente abaixo de `previous_cover_cell` → retorna `['ABC', 'DEF', 'GHI']`

### `TestCopyValues`
- copia `value` da célula origem para a célula destino no sheet mock (openpyxl: `.cell(row=..., column=...)`)

### `TestReorderDescriptionCells`
- chama `copy_values` exatamente 39 vezes (13 linhas × 3 colunas)
- primeira chamada: linha 18 → 17, coluna 1
- última chamada: linha 30 → 29, coluna 3

---

## `test_integration.py` — Integração com filesystem e Excel

`TestMoveFiles` e `TestCreateZip` rodam em qualquer plataforma (sem dependência de Excel).
`TestCreateExcelGrd` pula automaticamente fora do Windows ou se a fixture não existir (`@skipUnless(os.path.exists(FIXTURE_PATH) and sys.platform == 'win32', ...)`).

### `TestMoveFiles`
- arquivo com `emit=True` → movido para `3_Emitidos/<pasta-do-documento>/` e `msgbox` chamado uma vez
- arquivo com `emit=False` → permanece no local original, não é movido

### `TestCreateZip`
- arquivos com `emit=True` → incluídos no `.zip` criado no cwd com o nome da GRD
- arquivo com `emit=False` → não incluído no `.zip`
- lista vazia de emitidos → `.zip` criado mas vazio

### `TestCreateExcelGrd` *(requer Windows + `fixtures/IFS-XXXX-XXX-X-LD-XXXX.xlsx`)*

A fixture é um arquivo `.xlsx` sem macros usado como template de entrada. Os arquivos gerados também são `.xlsx`.

- `test_cria_arquivo_xlsx_na_pasta_ld`: copia o template, chama `create_excel_grd` com `ld_rev=-1` e `ld_information["ld_name"] = "IFS-2227-001-GER-LD-00001"` → verifica que `IFS-2227-001-GER-LD-00001_R0.xlsx` foi criado
- `test_cria_arquivo_xlsx_segunda_emissao`: encadeia duas chamadas — 1ª cria o R0 a partir da fixture, 2ª usa o R0 como base (`ld_rev=0`, `grd_number=2`) → verifica que `IFS-2227-001-GER-LD-00001_R1.xlsx` foi criado
