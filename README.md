# Validador

Script automatizado para validação de colaboradores no sistema web, com integração de planilha Excel para controle de processamento.

## 📋 Descrição

Este script automatiza o processo de validação de colaboradores em sistema web, realizando:

- Login no sistema
- Navegação até o módulo "Colaborador Avançado"
- Preenchimento de CPF
- Marcação de checkboxes de validação
- Gravação dos dados
- Controle de status na planilha Excel

## 🔧 Requisitos

### Python

- Python 3.11 ou 3.12 (recomendado)
- **Evite Python 3.14** (problemas de compatibilidade com urllib3)

### Dependências

```bash
pip install selenium==4.39.0
pip install undetected-chromedriver==3.5.5
pip install openpyxl==3.1.5
pip install urllib3==2.2.3
```

### Navegador

- Google Chrome instalado e atualizado

## 📁 Estrutura da Planilha

A planilha Excel deve ter a seguinte estrutura:

| Coluna | Conteúdo | Descrição                                        |
| ------ | --------- | -------------------------------------------------- |
| B      | CPF       | CPF do colaborador (com ou sem formatação)       |
| F      | Status    | Status do processamento (vazio, "Feito" ou "Erro") |

### Exemplo:

```
| A | B           | C | D | E | F      |
|---|-------------|---|---|---|--------|
| 1 | CPF         | . | . | . | Status |
| 2 | 123.456.789-00 | . | . | . |        |
| 3 | 987.654.321-00 | . | . | . | Feito  |
| 4 | 111.222.333-44 | . | . | . | Erro   |
```

## ⚙️ Configuração

Edite as constantes no início do arquivo `Validador.py`:

```python
LOGIN_URL = "https://seu-sistema.com.br/login"
PLANILHA_PATH = r"C:\Users\SEU_USUARIO\Downloads\SUA_PLANILHA.xlsx"
```

## 🚀 Como Usar

### 1. Prepare a Planilha

- Coloque os CPFs na **coluna B** (começando da linha 2)
- Certifique-se de que a **coluna F** existe (para status)
- Feche a planilha no Excel antes de executar o script

### 2. Execute o Script

```bash
python Validador.py
```

### 3. Processo Manual

O script abrirá o navegador e você deve:

1. Fazer login manualmente no sistema Cebraspe
2. Navegar até o módulo **SinCad** → **Colaborador** → **Colaborador Avançado**
3. Pressionar **Enter** no terminal quando estiver pronto

### 4. Processamento Automático

O script irá:

- Processar automaticamente cada CPF pendente
- Marcar "Feito" ou "Erro" na coluna F
- Salvar a planilha **a cada 10 registros processados**
- Exibir resumo ao final

## 📊 Lógica de Processamento

### O que é processado:

✅ Linhas com CPF na coluna B **E** status vazio na coluna F

### O que é pulado:

❌ Linhas com status "Feito" na coluna F
❌ Linhas com status "Erro" na coluna F
❌ Linhas sem CPF na coluna B

## 🔄 Fluxo de Validação

Para cada CPF, o script:

1. Preenche o campo CPF
2. Marca checkbox "Visualizar todas as Cidades"
3. Clica em "Pesquisar"
4. Aguarda 6 segundos para resultado carregar
5. Marca checkbox "Validade na Receita Federal?"
6. Clica em "Gravar"
7. Aceita alertas/popups de confirmação
8. Clica em "Voltar" para retornar à tela de pesquisa
9. Atualiza status na planilha

## 💾 Sistema de Salvamento

- **Salvamento parcial**: a cada 10 CPFs processados
- **Salvamento final**: ao concluir todos os CPFs
- **Status gravados**:
  - `Feito` - CPF validado com sucesso
  - `Erro` - Falha na validação

## ⚠️ Tratamento de Erros

O script continua executando mesmo se houver erros individuais:

- Erros são logados no console
- CPF com erro recebe status "Erro" na planilha
- Processamento continua para próximos CPFs

## 📝 Logs

O script exibe logs detalhados:

```
[PLANILHA] Carregando planilha: caminho/planilha.xlsx
[PLANILHA] Encontrados 4133 CPFs para processar.
[NAVEGAÇÃO] Abrindo página de login...
[VALIDAÇÃO] Processando 1/4133
[VALIDAÇÃO] Linha 26 - CPF 01226556213
[VALIDAÇÃO] ✓ CPF 01226556213 validado com sucesso!
[PLANILHA] Salvamento parcial após 10 atualizações (última linha 123).
```

## 🛑 Interrupção

Para interromper o script:

- Pressione `Ctrl+C` no terminal
- A planilha será salva automaticamente antes de fechar

## 🐛 Troubleshooting

### Erro: "Planilha não encontrada"

- Verifique o caminho em `PLANILHA_PATH`
- Use `r"caminho"` para evitar problemas com barras invertidas

### Erro: "Não foi possível abrir a planilha"

- Feche a planilha no Excel antes de executar
- Verifique permissões de escrita no arquivo

### Botão não encontrado (Pesquisar/Gravar/Voltar)

- Aumente os tempos de `time.sleep()` no código
- Verifique se está na página correta do sistema

### Import travando (Python 3.14)

- Desinstale Python 3.14
- Instale Python 3.12 ou 3.11
- Recrie o ambiente virtual

## 📦 Instalação Completa

```bash
# Clone ou baixe os arquivos
cd C:\Users\seu_usuario\Documents

# Crie ambiente virtual
python -m venv .venv

# Ative o ambiente
.venv\Scripts\Activate.ps1

# Instale dependências
pip install selenium==4.39.0 undetected-chromedriver==3.5.5 openpyxl==3.1.5 urllib3==2.2.3

# Execute
python Validador.py
```

## 📄 Licença

Uso interno

## 👤 Autor

Victor Vasconcelos

---

**Última atualização**: 29 de dezembro de 2025
