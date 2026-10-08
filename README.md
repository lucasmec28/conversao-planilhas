# Gerador de Tabelas - Robot Structural Analysis

Aplicativo Streamlit para transformar exportações do Autodesk Robot Structural Analysis em tabelas padronizadas para Memórias de Cálculo da Blossom.

## Entradas

O aplicativo utiliza três arquivos:

1. **CSV de combinações** exportado diretamente do Robot.
2. **CSV de reações** exportado diretamente do Robot.
3. **Excel ELU/ELS** criado em branco com o conteúdo da janela **Verificação de membro (ELS; ELU)** copiado e colado a partir de `A1`.

O cabeçalho do terceiro arquivo não é obrigatório.

## Ponto central: casos de carregamento variáveis

O aplicativo **não presume números fixos** para Peso Próprio, Carga Permanente, Sobrecarga, Temperatura etc.

Depois dos uploads, ele identifica os números de casos encontrados e apresenta uma tabela editável para o usuário informar:

- nome do carregamento;
- abreviação a usar nas combinações;
- se o caso deve ser incluído no quadro de reações.

O processamento só é liberado após a confirmação explícita desse mapeamento.

## Saídas

- Excel `.xlsx` com resumo, combinações ELU/ELS, Tabelas 8/9/10, validações e dados brutos ocultos para rastreabilidade;
- Word `.docx` com tabelas nativas, fonte **Times New Roman**, cabeçalhos repetidos e páginas em paisagem nas tabelas largas;
- pacote `.zip` contendo Excel + Word.

## Validações

O aplicativo informa, entre outros pontos:

- quantidade de combinações ELU e ELS reconhecidas;
- quantidade de membros lidos;
- quantidade de membros com resultados ELS;
- quantidade de nós de apoio;
- casos sem abreviação;
- membros sem nome no Robot (renomeados como `Barra_N`);
- índices ELU acima de 1,00;
- esbeltez acima do limite configurado;
- razões ELS acima de 1,00;
- combinações cuja sintaxe não pôde ser traduzida automaticamente.

## Execução local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy no Streamlit Community Cloud

1. Crie um repositório no GitHub e envie todos os arquivos desta pasta.
2. No Streamlit Community Cloud, crie um novo app apontando para o repositório.
3. Main file: `app.py`.
4. Recomenda-se Python 3.11 ou superior.
5. Não são necessários secrets.

## Configurações avançadas

A interface mantém escondidas em um `expander` as opções menos usadas:

- limite de esbeltez padrão;
- fator de conversão kgf → kN;
- fator de conversão kgfm → kN·m;
- manutenção dos dados brutos no Excel.

## Observações de engenharia

- O limite de esbeltez é global na V1. Para membros exclusivamente tracionados que devam usar outro limite, a evolução natural é permitir regra por grupo/membro.
- A Tabela 9 usa as proporções ELS fornecidas pelo próprio Robot; membros sem qualquer resultado ELS são omitidos.
- As reações exportadas no quadro final são apenas dos casos básicos selecionados pelo usuário; combinações marcadas pelo Robot com `(C)` não entram no quadro nominal.


## Formatação de saída (v1.2)
- Fonte: Times New Roman.
- Todas as páginas do Word são A4 retrato.
- As tabelas são ajustadas automaticamente a 100% da largura útil da página.
- Células categóricas repetidas em sequência são mescladas quando isso preserva o significado da tabela.
- Resultados numéricos, casos críticos e status não são mesclados.


## Ajustes v1.2
- Campo **Grupo** voltou ao mapeamento e permanece editável por caso.
- A Tabela 10 usa o grupo informado, permitindo agrupar casos mutuamente exclusivos.
- Mesclagem aplicada **somente** à Tabela 10 (reações): Base/nó e Grupo, sempre em valores iguais e sequenciais.
- Forças, momentos, carregamentos, casos e demais tabelas não são mesclados.
