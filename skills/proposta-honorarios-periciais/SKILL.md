---
name: proposta-honorarios-periciais
description: Elabora proposta de honorários periciais contábeis em Word, no layout tipográfico próprio do perito (NBC PP 01 R2, CPC 465), a partir do Mapa do Processo; usar logo após o mapa pericial ou quando pedirem orçamento/proposta de honorários.
---

# Proposta de Honorários Periciais (método CLARA — fase 1)

Ativação: imediatamente após a skill `mapa-do-processo-pericial`, ou quando o perito pedir "proposta de honorários", "orçamento da perícia", "estimativa de horas".
Saída única: `Proposta_Honorarios_<nº processo>.docx` gerado sobre o MODELO TIPOGRÁFICO DO PERITO + quadro-resumo no chat (máx. 10 linhas).

## 0. Regras de economia de tokens (obrigatórias)
- Fonte primária = Mapa do Processo já elaborado. NÃO reler os autos; consultar PDF só para lacuna específica (página pontual).
- Não transcrever o mapa nem a proposta no chat. Não explicar a norma ao usuário.
- Não redigir o documento: preencher o modelo com `scripts/gerar_proposta.py` a partir de um JSON. Correções = alterar o JSON e regerar.
- Perguntas: no máximo 1 rodada, só para o que impede a fixação do valor (papel do perito; valor-hora + link quando o processo não for do TJRS — §3). Demais lacunas → `[A CONFIRMAR]` no JSON.

## 1. Modelo tipográfico (obrigatório — não alterar)
Repositório: `https://github.com/ivanneryseidelperito-code/mapa-do-processo-pericial`, pasta `proposta-honorarios-periciais/`.
```
git clone --depth 1 https://github.com/ivanneryseidelperito-code/mapa-do-processo-pericial /tmp/mpp
python3 /tmp/mpp/proposta-honorarios-periciais/scripts/gerar_proposta.py dados.json Proposta_Honorarios_<proc>.docx
```
- `modelo/modelo_proposta_honorarios.docx` preserva o layout do perito: A4; margens sup. 3 cm, esq. 3 cm, inf. 2 cm, dir. 2 cm; cabeçalho com logotipo "Ivan Seidel — Perito Contábil e Financeiro"; rodapé verde com contatos e paginação "Pág. x-y"; corpo em Dupincel Text, títulos de seção em Segoe UI negrito 13 pt; quadro de partes em negrito itálico; caixa cinza com valor e extenso; tabela de etapas em 10 pt; alíneas a), b)… automáticas; controles de conteúdo vinculados às propriedades do documento (processo, partes, cidade).
- Estrutura fixa do modelo (aderente à NBC PP 01 R2): Endereçamento → Processo/Autor/Réu → Qualificação do perito → I – Da Nomeação (aceite, item 19; inexistência de impedimento/suspeição, CPC 144, 145 e 148 e itens 14 e 16; juntada da Certidão de Regularidade Profissional, item 5) → II – Dos Honorários (17 fatores do item 32; alíneas a–d; fonte do valor-hora, item 33; caixa de valor; alínea d: planejamento até a homologação, item 34, com um parágrafo por etapa — "Etapa n – título:" em negrito + atividades — seguido do quadro-resumo Descrição da tarefa | Estimativa do trabalho (em horas) | Valor hora adotado (R$) * | Valor dos serviços; nota de rodapé do valor-hora) → ressalvas de quesitos suplementares (CPC 469; item 37) e de despesas (item 35) → prazo estimado (CPC 465 e 476; itens 24 e 27) → III – Dos Pedidos (a: homologação após intimação das partes, CPC 465 §3º; b: depósito, CPC 82 e 95; c: levantamento de 50%, CPC 465 §4º e item 38; d: fixação do prazo) → fecho, local/data, assinatura.
- Proibido: recriar o documento com python-docx/docx-js, trocar fontes, margens, cabeçalho, rodapé ou ordem das seções. Se o repositório estiver inacessível, solicitar ao perito o anexo do modelo e usar `--modelo <arquivo>`.
- Se o perito enviar nova versão do seu layout: gerar novo modelo com marcadores (mesmos nomes `{{...}}`), substituir no repositório e registrar no §8.

## 2. Contexto (C) — extrair do mapa para o JSON
`juizo`, `cidade`, `processo`, `parte1_qualif`/`parte1_nome`, `parte2_qualif`/`parte2_nome`, `id_nomeacao` (Id. e página da decisão), `objeto` (complemento de "aspectos técnicos relativos a …"), `cpf` (somente se informado pelo perito; ausente → `[CPF]`), `prazo_dias` (prazo estimado para o laudo, compatível com as horas; ausente → `[A CONFIRMAR]`), `data` (opcional). Para dimensionar horas: nº de folhas, quesitos (juízo/autor/réu), meses do período, contratos/competências, documentos faltantes, diligências externas, justiça gratuita.
`etapas`: lista `[{descricao, horas, atividades}]` (ver §4).
Exemplo completo: `exemplo_dados.json`.

## 3. Valor-hora (VH) — padrão por tribunal
- **TJRS (CNJ com segmento `.8.21.`): VH padrão R$ 900,00 — FECONTABIL — https://www.fecontabil.org.br/pdfs/Tabela_de_Honorarios.pdf** (definido pelo perito em 25/11/2026). Usar sem perguntar, salvo se o perito informar outro valor/fonte na proposta em curso.
- JSON padrão TJRS: `"vh": {"valor": 900.00, "fonte_nome": "FECONTABIL", "fonte_url": "https://www.fecontabil.org.br/pdfs/Tabela_de_Honorarios.pdf", "nota": "Valor da hora técnica de referência divulgado na Tabela de Honorários da FECONTABIL."}`.
- **Demais tribunais:** o VH padrão NÃO se aplica. O perito informa (a) o valor-hora e (b) o link da fonte; sem ambos, não gerar a proposta (o script recusa a execução).
- Identificar o tribunal pelo nº CNJ (J.TR: 8.21 = TJRS). Padrões só mudam por instrução expressa do perito (atualizar esta seção e o §8).
- Não pesquisar, estimar ou atualizar o VH por conta própria. Conferir o link (WebFetch) somente se o perito solicitar.
- O link é citado no corpo (alínea c) e o valor/base na nota de rodapé — evidenciação dos critérios de formação do preço (NBC PP 01 R2, item 33).

## 4. Análise (A) — estimativa de horas
Horas por etapa = base + acréscimos (parâmetros calibráveis pelo §7; não normativos):

| Etapa | Base (h) | Acréscimos |
|---|---|---|
| E1 Exame dos autos e planejamento | 3 | +1 h/200 fls acima de 300 |
| E2 Diligências e solicitação de documentos (CPC 473 §3º) | 2 | +2 h/diligência externa; +1 h se documentos faltantes |
| E3 Análise documental e conferência | 4 | +1 h/12 meses; +1 h/contrato adicional |
| E4 Cálculos e planilhas | 6 | +1 h/12 competências; +2 h/cenário alternativo; +2 h se índices/juros com marco legal múltiplo |
| E5 Laudo e respostas aos quesitos (CPC 473) | 4 | +0,5 h/quesito |
| E6 Revisão técnica | 2 | +10% de (E3+E4+E5) |
| E7 Esclarecimentos (CPC 477 §2º) — reserva | 3 | +1 h se quesitos > 20 |

Fator de complexidade FC (NBC PP 01 R2, item 32 — relevância, vulto, risco, complexidade): 1,00 | 1,15 | 1,30 | 1,50. Horas de cada linha × FC, arredondadas a 0,5 h.
Planejamento (elaborado pela skill): cada etapa com horas > 0 gera, no JSON, um item `{descricao, horas, atividades}`, que o script converte em (i) parágrafo de planejamento e (ii) linha do quadro-resumo, na mesma ordem. Etapa sem horas → omitida (renumeração automática).
`atividades` = 1 a 2 frases, impessoais, específicas do caso (documentos, período, contratos, critérios e quesitos extraídos do mapa), partindo das bases abaixo; não copiar a base sem adaptá-la:
- E1 Exame dos autos e planejamento: leitura da inicial, contestação, decisões e quesitos; delimitação do objeto e dos procedimentos.
- E2 Diligências e solicitação de documentos: termo de diligência para requisição dos documentos faltantes (CPC 473 §3º), com comunicação prévia aos assistentes técnicos (NBC PP 01 R2, item 29).
- E3 Análise documental: exame de contratos, extratos, demonstrativos; conferência de consistência dos dados.
- E4 Cálculos e planilhas: planilhas de recálculo/apuração segundo os critérios fixados nos autos; cenários alternativos, se houver.
- E5 Elaboração do laudo e respostas aos quesitos: redação nos termos do CPC 473 e resposta fundamentada a cada quesito.
- E6 Revisão técnica: revisão integral de cálculos, respostas e fundamentação.
- E7 Esclarecimentos: reserva técnica para esclarecimentos (CPC 477 §2º).
Agrupar etapas (ex.: E5+E6) apenas por instrução do perito.
`valor_final` (opcional): informado pelo perito para arredondamento/desconto; o script gera a linha "Desconto/arredondamento". Sem ele, total = subtotal e a linha é suprimida.
Justiça gratuita: observar os limites da Res. CNJ 232/2016 (estadual), Res. CJF 305/2014 (federal) ou CLT art. 790-B (trabalho) e sinalizar ao perito antes de gerar.
Papel ATJ/ATE: o modelo é de perito do juízo; para assistente técnico ou trabalho extrajudicial, gerar com o mesmo modelo e sinalizar `[A CONFIRMAR]` nas seções I e III (contrato de prestação de serviços — NBC PP 01 R2, item 36).

## 5. Ajuste (A) — checagem antes de entregar
- O script valida: VH com fonte, ausência de marcadores não preenchidos, subtotal ≥ valor final, extenso. Não refazer contas manualmente.
- Converter para PDF (`soffice --headless --convert-to pdf`) e inspecionar visualmente a página 1 e a tabela: quebras, alinhamentos, linha de local/data.
- Fundamentos citados existentes e correspondentes; NBC PP 01 (R2) de 20/02/2025 (DOU 14/03/2025).
- Nenhum `[A CONFIRMAR]`/`[CPF]` oculto: listar no resumo.

## 6. Entrega
Enviar o .docx com SendUserFile. No chat: linhas da tabela (horas), FC, VH + fonte, subtotal, desconto, total, prazo, pendências — incluindo, sempre, o lembrete de anexar a Certidão de Regularidade Profissional atualizada (NBC PP 01 R2, item 5). Nada além.

## 7. Protocolo de autoaprendizado
Acionado quando o perito criticar, corrigir ou informar resultado (ex.: "horas de cálculo ficaram baixas", "o juiz reduziu", "gastei 30 h").
1. Corrigir o documento atual (JSON → regerar).
2. Converter a crítica em regra geral, verificável e curta (1 linha), indicando a seção afetada.
3. Horas realizadas × estimadas: ajustar o parâmetro da etapa por média ponderada (70% atual, 30% observado).
4. Chamar `propose_skills` (kind: improvement, target: proposta-honorarios-periciais) com `edits` que alterem a regra/parâmetro e acrescentem linha ao §8. O perito aprova no cartão; sem aprovação, nada é alterado.
5. Após aprovação, se o perito pedir, sincronizar a SKILL.md no repositório GitHub (commit na pasta `proposta-honorarios-periciais/`).
6. Não registrar dados identificáveis das partes; apenas natureza da ação e aprendizado.

## 8. Registro de Aprendizado
Formato: `RA-nn | data | natureza da ação | crítica | regra adotada | seção alterada`
- RA-01 | 25/11/2026 | geral | proposta deve seguir o layout tipográfico próprio do perito | geração exclusiva por preenchimento do modelo (§1) | §1, §0
- RA-02 | 25/11/2026 | geral | valor-hora e respectiva fonte variam por proposta | VH e link informados pelo perito a cada proposta; vedada estimativa própria (§3) | §3
- RA-03 | 25/11/2026 | geral | modelo deve contemplar todos os tópicos da NBC PP 01 (R2) aplicáveis à proposta | inclusão no modelo de: impedimento/suspeição, certidão de regularidade, 17 fatores do item 32, planejamento até a homologação, complementação justificada de quesitos, prazo e pedidos correspondentes | §1, §2, §6
- RA-04 | 25/11/2026 | geral | planejamento deve ser elaborado pela skill, etapa por etapa, antes do quadro-resumo | parágrafo por etapa (título + atividades específicas do caso) e quadro com uma linha por etapa; colunas renomeadas | §1, §2, §4
- RA-05 | 25/11/2026 | geral | valor-hora não deve ser informado a cada proposta, mas o padrão vale só para o TJRS | VH padrão R$ 900,00 (FECONTABIL) apenas para processos do TJRS (8.21); demais tribunais seguem exigindo VH + link | §0, §3
