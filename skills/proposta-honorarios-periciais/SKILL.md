---
name: proposta-honorarios-periciais
description: Elabora proposta de honorários periciais contábeis em Word, no layout tipográfico próprio do perito (NBC PP 01 R2, CPC 465), a partir do Mapa do Processo; usar logo após o mapa pericial ou quando pedirem orçamento/proposta de honorários.
---

# Proposta de Honorários Periciais (método CLARA — fase 1)

Ativação: logo após `mapa-do-processo-pericial`, ou pedido de "proposta de honorários", "orçamento da perícia", "estimativa de horas".
Saída: `Proposta_Honorarios_<nº processo>.docx` gerado sobre o modelo do perito + resumo no chat (máx. 10 linhas).

## 0. Economia de tokens (obrigatória)
- Fonte = Mapa do Processo. Não reler os autos; PDF só para lacuna pontual.
- Não transcrever mapa nem proposta no chat; não explicar normas.
- Não redigir: preencher o modelo via `scripts/gerar_proposta.py` a partir de JSON. Correção = alterar JSON e regerar.
- Perguntas: 1 rodada, só o que impede fixar o valor (papel do perito; VH + link fora do TJRS — §3). Demais lacunas → `[A CONFIRMAR]`.

## 1. Modelo (obrigatório — não alterar)
```
git clone --depth 1 https://github.com/ivanneryseidelperito-code/mapa-do-processo-pericial /tmp/mpp
python3 /tmp/mpp/proposta-honorarios-periciais/scripts/gerar_proposta.py dados.json Proposta_Honorarios_<proc>.docx
```
- `modelo/modelo_proposta_honorarios.docx` já contém layout e estrutura fixa aderente à NBC PP 01 (R2) (nomeação, honorários, ressalvas, prazo, pedidos). Detalhes: README do repositório.
- Proibido recriar o documento (python-docx/docx-js), trocar fontes, margens, cabeçalho, rodapé ou ordem das seções. Repositório inacessível → pedir o modelo ao perito e usar `--modelo <arquivo>`.
- Novo layout do perito: gerar modelo com os mesmos marcadores `{{...}}`, substituir no repositório e registrar em `APRENDIZADO.md`.

## 2. Contexto — campos do JSON (extrair do mapa)
`juizo`, `cidade`, `processo`, `parte1_qualif`/`parte1_nome`, `parte2_qualif`/`parte2_nome`, `id_nomeacao` (Id. e página), `objeto` (complemento de "aspectos técnicos relativos a …"), `cpf` (só se informado; senão `[CPF]`), `prazo_dias` (compatível com as horas; senão `[A CONFIRMAR]`), `data` (opcional), `vh` (§3), `etapas` (§4), `valor_final` (opcional). Exemplo: `exemplo_dados.json`.
Dimensionamento: nº de folhas, quesitos (juízo/autor/réu), meses do período, contratos/competências, documentos faltantes, diligências externas, justiça gratuita.

## 3. Valor-hora (VH)
- **TJRS (CNJ `.8.21.`): R$ 900,00 — FECONTABIL** (definido pelo perito em 25/11/2026), sem perguntar, salvo outro valor informado na proposta em curso:
  `"vh": {"valor": 900.00, "fonte_nome": "FECONTABIL", "fonte_url": "https://www.fecontabil.org.br/pdfs/Tabela_de_Honorarios.pdf", "nota": "Valor da hora técnica de referência divulgado na Tabela de Honorários da FECONTABIL."}`
- **Demais tribunais:** perito informa valor e link da fonte; sem ambos, não gerar (o script recusa).
- Não pesquisar nem estimar VH. Conferir o link só se o perito pedir. Padrões mudam só por instrução expressa (atualizar este § e `APRENDIZADO.md`).

## 4. Análise — horas por etapa
Horas = base + acréscimos (parâmetros calibráveis, não normativos):

| Etapa | Base (h) | Acréscimos |
|---|---|---|
| E1 Exame dos autos e planejamento | 3 | +1 h/200 fls acima de 300 |
| E2 Diligências e solicitação de documentos | 2 | +2 h/diligência externa; +1 h se faltam documentos |
| E3 Análise documental e conferência | 4 | +1 h/12 meses; +1 h/contrato adicional |
| E4 Cálculos e planilhas | 6 | +1 h/12 competências; +2 h/cenário alternativo; +2 h se marco legal múltiplo de índices/juros |
| E5 Laudo e respostas aos quesitos | 4 | +0,5 h/quesito |
| E6 Revisão técnica | 2 | +10% de (E3+E4+E5) |
| E7 Esclarecimentos (CPC 477 §2º) — reserva | 3 | +1 h se quesitos > 20 |

Fator de complexidade FC (item 32): 1,00 | 1,15 | 1,30 | 1,50. Horas × FC, arredondar a 0,5 h. Etapa com 0 h é omitida (renumeração automática).
Cada etapa → `{descricao, horas, atividades}`; o script gera o parágrafo de planejamento e a linha do quadro, na mesma ordem.
`atividades`: 1–2 frases impessoais, específicas do caso (documentos, período, contratos, critérios, quesitos), adaptando a base:
- E1 leitura de inicial, contestação, decisões e quesitos; delimitação do objeto e procedimentos.
- E2 termo de diligência para documentos faltantes (CPC 473 §3º), com comunicação prévia aos assistentes (item 29).
- E3 exame de contratos, extratos, demonstrativos; consistência dos dados.
- E4 planilhas de recálculo/apuração pelos critérios dos autos; cenários alternativos, se houver.
- E5 redação nos termos do CPC 473 e resposta fundamentada a cada quesito.
- E6 revisão integral de cálculos, respostas e fundamentação.
- E7 reserva para esclarecimentos (CPC 477 §2º).
Agrupar etapas só por instrução do perito. `valor_final` gera a linha "Desconto/arredondamento"; sem ele, total = subtotal.
Justiça gratuita: limites da Res. CNJ 232/2016 (estadual), Res. CJF 305/2014 (federal) ou CLT 790-B; sinalizar ao perito antes de gerar.
ATJ/ATE: mesmo modelo, com `[A CONFIRMAR]` nas seções I e III (contrato de serviços — item 36).

## 5. Checagem
- O script valida VH com fonte, marcadores vazios, subtotal ≥ valor final e extenso; não refazer contas.
- `soffice --headless --convert-to pdf` e inspecionar a página 1 e a tabela (quebras, alinhamentos, local/data).
- Citar a NBC PP 01 (R2) de 20/02/2025 (DOU 14/03/2025). Listar no resumo todo `[A CONFIRMAR]`/`[CPF]`.

## 6. Entrega
SendUserFile com o .docx. No chat, só: horas por etapa, FC, VH + fonte, subtotal, desconto, total, prazo, pendências e o lembrete da Certidão de Regularidade Profissional atualizada (item 5).

## 7. Autoaprendizado
Quando o perito criticar, corrigir ou informar resultado ("o juiz reduziu", "gastei 30 h"):
1. Corrigir o documento atual (JSON → regerar).
2. Converter a crítica em regra curta e verificável, indicando a seção.
3. Horas realizadas × estimadas: novo parâmetro = 70% atual + 30% observado.
4. `propose_skills` (kind: improvement, target: proposta-honorarios-periciais) com a alteração; o perito aprova no cartão.
5. Registrar em `APRENDIZADO.md` (formato lá descrito), sem dados identificáveis das partes.
6. Após aprovação, se o perito pedir, sincronizar com o GitHub.
