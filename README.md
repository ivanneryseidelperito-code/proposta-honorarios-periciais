# proposta-honorarios-periciais

Plugin do Claude com a skill "Proposta de Honorários Periciais" (Ivan Nery Seidel – Perito Contábil), acionada após o Mapa do Processo.

- `.claude-plugin/plugin.json` – manifesto do plugin
- `skills/proposta-honorarios-periciais/SKILL.md` – instruções da skill
- `skills/proposta-honorarios-periciais/modelo/modelo_proposta_honorarios.docx` – layout tipográfico do perito, com marcadores `{{...}}`
- `skills/proposta-honorarios-periciais/scripts/gerar_proposta.py` – preenche o modelo a partir de `dados.json`
- `skills/proposta-honorarios-periciais/exemplo_dados.json` – exemplo de entrada

```
python3 skills/proposta-honorarios-periciais/scripts/gerar_proposta.py dados.json Proposta.docx
```
