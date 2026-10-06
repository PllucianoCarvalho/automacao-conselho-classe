<div align="center">

# 📊 Automação de Conselhos de Classe

### Sistema para consolidação automatizada de dados pedagógicos e geração de relatórios

<br>

<img src="https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
<img src="https://img.shields.io/badge/Pandas-Data%20Processing-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas">
<img src="https://img.shields.io/badge/Excel-Automation-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="Excel">
<img src="https://img.shields.io/badge/Google%20Colab-Development-F9AB00?style=for-the-badge&logo=googlecolab&logoColor=white" alt="Google Colab">

</div>

---

## 🎯 Sobre o projeto

Este projeto foi desenvolvido para **automatizar o processamento e a consolidação de dados pedagógicos utilizados em conselhos de classe**.

O sistema integra informações provenientes de múltiplos arquivos Excel, realiza validações e tratamentos dos dados, consolida os registros dos estudantes e gera automaticamente **relatórios pedagógicos padronizados em documentos editáveis**.

A solução foi desenvolvida a partir de uma necessidade real do ambiente escolar, buscando reduzir etapas manuais, agilizar a preparação dos conselhos de classe e tornar as informações mais **concisas, organizadas e padronizadas**.

---

## 🚀 Problema

A preparação dos dados para os conselhos de classe envolvia o processamento manual de diversas planilhas produzidas pelos professores.

Esse processo apresentava alguns desafios:

- 📁 grande quantidade de arquivos;
- 🔄 necessidade de consolidação de informações;
- 🧹 tratamento de dados inconsistentes;
- 👤 registros distribuídos entre diferentes planilhas;
- 📝 elaboração manual dos relatórios;
- ⏱️ elevado tempo de preparação;
- 📄 dificuldade em manter um padrão de apresentação.

O projeto surgiu como uma solução para **automatizar esse fluxo de trabalho**.

---

## 💡 Solução

O sistema automatiza o fluxo completo:

```text
	 📁 ARQUIVOS EXCEL
		 │
		 ▼
	 🔎 IDENTIFICAÇÃO
	  DAS TURMAS
		 │
		 ▼
	 🧹 VALIDAÇÃO E
	 PADRONIZAÇÃO
		 │
		 ▼
	🔗 CONSOLIDAÇÃO
	   DOS DADOS
		 │
		 ▼
	📊 AGREGAÇÃO DOS
	   INDICADORES
		 │
		 ▼
	📝 GERAÇÃO DOS
	   RELATÓRIOS
		 │
		 ▼
	 📄 DOCUMENTOS
	     .DOCX
```

---

## ⚙️ Funcionalidades

### 📥 Processamento de arquivos

- Leitura de múltiplos arquivos Excel;
- identificação das abas correspondentes às turmas;
- processamento automatizado dos dados;
- suporte a diferentes conjuntos de turmas.

### 🧹 Tratamento e validação

- Padronização dos nomes dos estudantes;
- identificação de dados ausentes;
- validação de valores numéricos;
- identificação de valores fora dos intervalos esperados;
- remoção de registros inválidos;
- registro de ocorrências em arquivo de log.

### 📊 Consolidação

O sistema consolida informações provenientes de diferentes registros do mesmo estudante.

Entre os dados processados estão:

- nível de proficiência;
- comprometimento com a aprendizagem;
- proatividade e curiosidade;
- participação;
- comportamento;
- assiduidade;
- existência de relatório;
- habilidades identificadas.

Os indicadores numéricos podem ser agregados automaticamente para produzir uma visão consolidada dos dados.

### 📝 Geração automática de relatórios

A partir dos dados consolidados, o sistema produz descrições textuais dos estudantes e gera documentos `.docx` organizados por turma.

| Nº | Aluno | Relatório |
| --- | --- | --- |
| 01 | Estudante | Relatório pedagógico consolidado |
| 02 | Estudante | Relatório pedagógico consolidado |
| 03 | Estudante | Relatório pedagógico consolidado |

---

## 🧠 Tecnologias utilizadas

| Tecnologia | Utilização |
| --- | --- |
| 🐍 **Python** | Desenvolvimento do sistema |
| 🐼 **Pandas** | Manipulação e consolidação dos dados |
| 📊 **OpenPyXL** | Processamento de arquivos Excel |
| 📄 **python-docx** | Geração dos documentos Word |
| ⏳ **tqdm** | Acompanhamento do processamento |
| 📝 **Logging** | Registro de eventos e erros |
| 💻 **Python local** | Execução do processamento no computador |

---

## 🏗️ Arquitetura do processamento

O sistema foi estruturado em etapas independentes para facilitar manutenção e evolução.

### 1. Entrada

Recebimento dos arquivos Excel produzidos pelos professores.

### 2. Identificação

Localização das abas correspondentes às turmas configuradas.

### 3. Validação

Verificação da estrutura dos dados e dos valores recebidos.

### 4. Normalização

Padronização dos nomes e tratamento de dados ausentes ou inconsistentes.

### 5. Consolidação

Agrupamento das informações de um mesmo estudante provenientes de diferentes arquivos.

### 6. Processamento

Cálculo e agregação dos indicadores pedagógicos.

### 7. Geração textual

Conversão dos indicadores em descrições pedagógicas estruturadas.

### 8. Exportação

Criação automática dos documentos finais em formato `.docx`.

---

## 📈 Impacto

A automação foi desenvolvida para reduzir significativamente o trabalho manual envolvido na preparação dos conselhos de classe.

### Resultados observados

- ⚡ maior velocidade no processamento;
- 📉 redução de tarefas repetitivas;
- 📊 maior padronização dos dados;
- 📝 relatórios mais concisos;
- 🔄 possibilidade de reutilização do processo;
- 🧩 facilidade de adaptação para diferentes turmas;
- 📄 geração automatizada dos documentos finais.

> O projeto transforma uma atividade predominantemente manual em um fluxo de processamento de dados estruturado e automatizado.

---

## 🎓 Aplicação educacional

Embora seja uma solução de automação de dados, o projeto foi desenvolvido especificamente para atender uma **necessidade do ambiente escolar**.

Sua aplicação envolve a utilização de recursos de programação e processamento de dados para apoiar processos de gestão pedagógica.

O projeto demonstra a aplicação prática de:

**Programação + Ciência de Dados + Automação + Tecnologia Educacional**

---

## 🔒 Privacidade

Os arquivos utilizados no processamento podem conter **dados pessoais e informações pedagógicas de estudantes**.

Por esse motivo:

- a planilha `PRE-CONSELHO 2025 _ GITHUB.xlsx` foi revisada manualmente pelo responsável e liberada para publicação;
- dados pessoais não devem ser publicados;
- recomenda-se utilizar dados fictícios ou anonimizados para testes;
- planilhas locais e documentos gerados durante a execução permanecem fora do controle de versão.

---

## 📂 Estrutura do projeto

```text
automacao-conselho-classe/
│
├── README.md
├── requirements.txt
├── PRE-CONSELHO 2025 _ GITHUB.xlsx  # planilha de demonstração sanitizada
├── src/
│   └── gerar_relatorios.py
├── dados/                 # arquivos Excel de entrada (não versionados)
└── relatorios_gerados/    # documentos DOCX gerados (não versionados)
```

---

## ▶️ Execução

### 1. Instalar as dependências

```bash
pip install -r requirements.txt
```

### 2. Executar o sistema

Para processar a planilha de demonstração que está na raiz do projeto:

```bash
python src/gerar_relatorios.py --entrada .
```

Para processar suas próprias planilhas, coloque os arquivos `.xlsx` em `dados/` e execute:

```bash
python src/gerar_relatorios.py
```

Os relatórios `.docx` serão salvos em `relatorios_gerados/`. As opções `--entrada` e `--saida` permitem escolher outras pastas:

```bash
python src/gerar_relatorios.py --entrada caminho/para/planilhas --saida caminho/para/relatorios
```

O script procura arquivos `.xlsx` diretamente na pasta de entrada, sem percorrer subpastas. Todos os arquivos encontrados são processados e consolidados juntos. Os relatórios existentes com o mesmo nome de turma são substituídos ao executar novamente.

---

## 📝 Como editar a planilha

O formato é posicional: o script ignora as três primeiras linhas de cada aba e interpreta cada linha a partir da quarta como um registro de estudante. Não insira uma linha de cabeçalho entre essas linhas iniciais e os dados. Mantenha as colunas nesta ordem:

| Posição | Campo | Regra |
| --- | --- | --- |
| 1 | Nº | Número ou identificador exibido no relatório |
| 2 | ALUNO | Nome usado para agrupar registros; não pode ficar vazio |
| 3 | NIVEL DE PROEFICIENCIA DO ALUNO | Nota numérica de 0 a 4 |
| 4 | COMPROMETIDO COM APRENDIZAGEM? | Nota numérica de 0 a 4 |
| 5 | PROATIVO E/OU CURIOSO? | Nota numérica de 0 a 4 |
| 6 | PARTICIPATIVO? | Nota numérica de 0 a 4 |
| 7 | EDUCADO(A)? | Nota numérica de 0 a 4 |
| 8 | FALTOSO? | Nota numérica de 0 a 4; a descrição de assiduidade segue a escala configurada no código |
| 9 | POSSUI RELATORIO? | Preencha com `SIM` ou `NÃO` |
| 10 | EM QUAL ÁREA APRESENTA HABILIDADE? | Texto; se houver várias habilidades, separe-as por vírgulas |
| 11 | NIVEL DO ALUNO SEGUNDO AS RESPOSTAS | Coluna preservada no formato, mas não utilizada no relatório atual |

As notas devem estar entre 0 e 4. Valores fora desse intervalo são descartados e registrados no log; células vazias são tratadas como ausência de informação. Colunas extras são ignoradas. Mantenha o nome da aba igual ao da turma configurada.

### Consolidação e limites

- Abas processadas nesta versão: `3A`, `6A`, `6B`, `6C`, `7A`, `7B`, `7C`, `8A`, `8B`, `8C` e `9A`. A aba `INICIAL` da planilha de demonstração não é processada.
- Apenas arquivos `.xlsx` são lidos; arquivos `.xls`, `.xlsm` e planilhas dentro de subpastas não são incluídos.
- Registros com o mesmo nome, após remover espaços nas extremidades e converter para maiúsculas, são tratados como o mesmo estudante dentro da turma. As notas são calculadas pela média; habilidades são reunidas e o indicador de relatório fica `SIM` se algum registro tiver `SIM`.
- O agrupamento usa somente o nome, não o campo `Nº`. Estudantes diferentes com nomes iguais podem ser consolidados por engano; confira a unicidade dos nomes antes de processar vários arquivos.
- Não há um limite fixo de arquivos ou estudantes definido no código. O processamento mantém os dados em memória, portanto o volume suportado depende dos recursos disponíveis; arquivos grandes não foram submetidos a teste de desempenho.
- A configuração de turmas, ordem das colunas e descrições das notas fica em `CONFIG` e `DESCRICOES`, no início de `src/gerar_relatorios.py`. Alterar a estrutura da planilha exige atualizar essa configuração e validar o resultado.

---

## 🔮 Possibilidades de evolução

Entre as possibilidades futuras estão:

- interface gráfica para execução sem necessidade de código;
- processamento automático de arquivos em uma pasta;
- exportação para PDF;
- geração de indicadores estatísticos;
- criação de dashboard para acompanhamento;
- configuração externa das turmas e critérios;
- integração com sistemas escolares;
- anonimização automática dos dados;
- criação de modelos de relatório configuráveis.

---

## 👨‍💻 Autor

**Paulo Luciano de Carvalho Maciel**

Professor de Física e Robótica<br>
Desenvolvimento de tecnologias educacionais

---

<div align="center">

### 💻 Tecnologia aplicada à educação
**Automação • Dados • Programação • Gestão Pedagógica**

</div>
