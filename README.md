# 📊 Agregador de Dados para Declaração de Imposto de Renda (`GREEN_IR.xlsx`) 
 
Ferramenta desenvolvida no **Excel** para centralizar, organizar e validar informações essenciais para a **Declaração de Imposto de Renda Pessoa Física**, reunindo dados do titular, informes de rendimentos bancários, lançamentos de entradas e saídas (com identificação de despesas dedutíveis) e tabelas de apoio. 
 
--- 
 
## 🎯 Objetivo do Projeto 
 
Criar um agregador de dados estruturado no Excel que permita controlar registros de forma organizada e validada, com: 
- Separação clara por contextos em abas dedicadas (`TITULAR`, `INFORMES`, `NOTAS` e `TABELAS`). 
- **Tabelas oficiais do Excel** para expansão dinâmica dos lançamentos e listas de apoio. 
- **Regras de validação de dados** para reduzir erros de digitação (CPF, data, e-mail, opções de lista e valores numéricos). 
- **Fórmulas dinâmicas** para consolidação de totais, saldo e despesas dedutíveis. 
 
--- 
 
## 🗂️ Estrutura da Planilha 
 
| Aba | Finalidade | Principais Elementos | 
| :--- | :--- | :--- | 
| **`TITULAR`** | Cadastro dos dados pessoais, endereço, contato e informações de declaração do contribuinte. | Campos de identificação e validações em `D8:D9`, `D17` e `D18:D20`. | 
| **`INFORMES`** | Registro dos informes de rendimentos bancários e referência de anexos/extratos. | Blocos por banco (ex.: Santander, PagBank, Nubank), totalizador em `D8` e validação em `D12`, `D17` e `D22`. | 
| **`NOTAS`** | Controle de entradas (receitas) e saídas (despesas) com resumo consolidado. | Tabelas `tbEntradas` (`C8:E25`) e `tbSaidas` (`G8:I25`), linhas de total (`E25` e `I25`) e painel de resumo (`H1:I5`). | 
| **`TABELAS`** | Base de dados de apoio para categorização e padronização. | Tabelas `tbBancos` (`A1:A51`), `tbCatEntradas` (`C1:C4`) e `tbCatSaidas` (`E1:F7`, incluindo coluna `DEDUTÍVEL`). | 
 
--- 
 
## 🧮 Fórmulas Utilizadas 
 
### Aba `INFORMES` 
- **Total de rendimentos / valores atuais informados (`D8`):** 
  ```excel 
  =SUMIF(C11:C23,"VALOR ATUAL",D11:D23) 
 

Soma os valores das linhas onde a coluna C corresponde exatamente a "VALOR ATUAL". 

Aba NOTAS 

Total dinâmico de entradas (E25): 

Excel 

=SUBTOTAL(109,tbEntradas[VALOR]) 
 

Calcula a soma (109) da coluna VALOR da tabela tbEntradas, respeitando filtros aplicados na tabela. 

Total dinâmico de saídas (I25): 

Excel 

=SUBTOTAL(109,tbSaidas[VALOR]) 
 

Calcula a soma (109) da coluna VALOR da tabela tbSaidas, respeitando filtros aplicados na tabela. 

Total de entradas no resumo (I2): =E25 

Total de saídas no resumo (I3): =I25 

Saldo entre entradas e saídas (I4): =I2-I3 

Total dedutível (I5): fórmula matricial (ArrayFormula) para somar os valores de tbSaidas[VALOR] cuja categoria em tbSaidas[CATEGORIA] está marcada como "Sim" na coluna DEDUTÍVEL (tbCatSaidas). 

✅ Validações de Dados Implementadas 

Aba TITULAR 

CPF (D8): número inteiro entre 1 e 99999999999, com mensagem de orientação e erro para evitar pontos ou traços. 

Data de nascimento (D9): data menor ou igual à data atual (TODAY()). 

E-mail (D17): validação personalizada verificando a presença de "@" e "." após o "@": 

Excel 

=AND(ISNUMBER(FIND("@",D17)),ISNUMBER(FIND(".",D17,FIND("@",D17)))) 
 

Campos "Sim" / "Não" (D18:D20): lista suspensa (Sim,Não) para alterações em relação à entrega anterior, dependente cônjuge e residente no exterior. 

Aba INFORMES 

Valores (D12, D17, D22): número decimal maior ou igual a 0. 

Aba NOTAS 

Datas (C9:C24 e G9:G24): validação de data no intervalo do ano-base (01/01/2026 a 31/12/2026). 

Valores (E9:E24 e I9:I24): número decimal maior que 0. 

📌 Observação sobre os Dados de Exemplo 

Os registros presentes na planilha servem como exemplo de preenchimento e estruturação da ferramenta e devem ser substituídos pelos dados reais do contribuinte no momento do uso.
