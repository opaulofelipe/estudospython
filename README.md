# Painel de Estudos

Dashboard em Python/Streamlit para acompanhar o progresso da planilha `Estudos.xlsx`.

## Estrutura esperada da planilha

A planilha pode ter qualquer nome de aba, mas a aba `estudos` é priorizada quando existir. As colunas obrigatórias são:

- `Disciplina`
- `Aulas totais`
- `Dia`
- `Instituição`

Você pode adicionar ou excluir disciplinas. Depois, use **Sincronizar planilha** no painel. O progresso é salvo separadamente em `data/progresso.json`, então a sincronização não depende da ordem das linhas da planilha.

A identidade de cada disciplina é formada pelo **nome da Disciplina**. Isso permite mudar dia ou instituição sem perder o progresso. Evite nomes de disciplina duplicados.

## Recursos

- duas abas: **Registrar aulas** e **Visão geral**;
- área de **Matérias do dia**;
- registro da próxima aula em um clique;
- grade numerada para marcar/desmarcar qualquer aula;
- ao marcar uma aula com anteriores pendentes, aparece uma confirmação para preencher todo o intervalo;
- concluir todas as aulas de uma disciplina;
- concluir o plano inteiro;
- sincronizar a planilha preservando progresso existente;
- progresso geral, gráfico de rosca, avanço por disciplina e ritmo semanal;
- backup do progresso em JSON;
- interface responsiva na paleta `#8C93A8`, `#62466B`, `#45364B`, `#2D2327`.

## Executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

O app procura `Estudos.xlsx` na mesma pasta de `app.py`.

## Observação sobre persistência em hospedagens gratuitas

O arquivo `data/progresso.json` persiste normalmente no computador/servidor enquanto o armazenamento local for permanente. Algumas plataformas gratuitas usam disco efêmero e podem apagar arquivos locais quando a aplicação reinicia ou é reconstruída. Nesses casos, use o botão de backup ou conecte depois uma persistência externa.
