# 📊 Sistema de Consolidação de Planilhas de Atendimentos

Sistema automatizado para consolidar múltiplas planilhas de atendimentos em uma única planilha mãe com métricas e relatórios.

---

## 🚀 Como Usar

### **1. Requisitos**

- Windows 7/8/10/11 (64-bit)
- **NADA MAIS!** (Não precisa Python, Excel, ou bibliotecas)

### **2. Estrutura de Arquivos**

```
Sua_Pasta/
├── atualizar_planilhas.exe  ← O programa
└── filhas/                   ← Suas planilhas aqui
    ├── NOME_SOBRENOME - ATENDIMENTOS - DD-MM-AA.xlsm
    ├── NOME_SOBRENOME - ATENDIMENTOS - DD-MM-AA.xlsm
    └── ...
```

### **3. Executar o Programa**

1. **Duplo clique** em `atualizar_planilhas.exe`

2. **Se o Excel estiver aberto:**

   ```
   💡 OPÇÕES:
      1. Fechar Excel automaticamente
      2. Criar arquivo temporário (abre em nova janela)
      3. Cancelar operação

   🤔 Escolha uma opção (1/2/3):
   ```

3. **Aguardar o processamento**

   - O programa lê todas as planilhas da pasta `filhas/`
   - Consolida os dados
   - Cria backup automático

4. **Resultado:**
   ```
   ✅ Processo concluído!
   🤔 Deseja abrir a planilha agora? (S/N):
   ```

### **4. Arquivos Gerados**

```
Sua_Pasta/
├── atualizar_planilhas.exe
├── filhas/
├── PLANILHA_MAE.xlsx        ← Planilha consolidada ✅
├── backup/
│   └── PLANILHA_MAE_BACKUP.xlsx
├── log_compilacao.txt        ← Log da execução
└── PLANILHA_TEMP_*.xlsx     ← Temporário (se usar opção 2)
```

---

## 📋 Planilha Consolidada (PLANILHA_MAE.xlsx)

### **Abas Geradas:**

1. **COMPILE GERAL**

   - Todos os atendimentos consolidados
   - Colunas extras: PRIMEIRO_NOME, DATA_ARQUIVO, ARQUIVO

2. **MÉTRICAS**

   - Atendimentos Totais
   - Atendimentos Finalizados
   - % Finalizados
   - Tempo Total e Médio
   - Tabelas por Setor e por Responsável

3. **Abas Individuais** (Amanda, Raphaela, etc.)
   - Dados filtrados por responsável
   - Mesmas colunas do arquivo original

---

## 📦 Como Distribuir

### **Método 1: Copiar Pasta Completa**

1. Copie a pasta inteira para:

   - Pendrive
   - Email (zipar antes)
   - Rede compartilhada
   - OneDrive/Google Drive

2. Usuário descompacta e executa

### **Método 2: Criar Pacote ZIP**

```
1. Selecione:
   - atualizar_planilhas.exe
   - filhas/ (com planilhas exemplo)

2. Botão direito → Enviar para → Pasta compactada

3. Envie o arquivo .zip
```

### **Método 3: Instalação em Rede**

```
\\Servidor\Compartilhado\Atendimentos\
├── atualizar_planilhas.exe
└── filhas/
```

Usuários acessam e executam direto da rede.

---

## 📝 Regras de Nomenclatura

### **Arquivos na pasta `filhas/`:**

✅ **Formato Correto:**

```
NOME_SOBRENOME - ATENDIMENTOS - DD-MM-AA.xlsm
NOME_SOBRENOME - ATENDIMENTOS - DD-MM-AAAA.xlsm
```

✅ **Exemplos Válidos:**

```
AMANDA_PINHEIRO - ATENDIMENTOS - 16-10-25.xlsm
RAPHAELA_MARQUES - ATENDIMENTOS_20-10-25.xlsm
JOAO_SILVA - ATENDIMENTOS - 22-10-2025.xlsx
```

❌ **Exemplos Inválidos:**

```
planilha atendimentos.xlsx  (sem padrão)
Amanda - 16-10-25.xlsm     (falta SOBRENOME e ATENDIMENTOS)
atendimentos_outubro.xlsm   (sem nome e data)
```

---

## 🔧 Funcionalidades Especiais

### **Modo Temporário (Opção 2)**

Quando o Excel está aberto:

- Cria arquivo `PLANILHA_TEMP_YYYYMMDD_HHMMSS.xlsx`
- Abre em nova janela do Excel
- Não interfere no arquivo principal
- Remove temporários antigos automaticamente

### **Backup Automático**

- Sempre cria backup antes de atualizar
- Mantém apenas o último backup
- Salvo em `backup/PLANILHA_MAE_BACKUP.xlsx`

### **Log de Execução**

Arquivo `log_compilacao.txt` contém:

```
Execução em 2025-10-22 17:30:45
Total de registros: 143
--------------------------------------------------
✅ OK: AMANDA_PINHEIRO - ATENDIMENTOS - 16-10-25.xlsm (29 linhas)
✅ OK: RAPHAELA_MARQUES - ATENDIMENTOS_20-10-25.xlsm (114 linhas)
```

---

## ⚠️ Solução de Problemas

### **"Diretório de filhas não encontrado"**

- Certifique-se que a pasta `filhas/` existe ao lado do .exe
- Verifique se há planilhas dentro da pasta

### **"Nome inválido"**

- Renomeie os arquivos seguindo o padrão: `NOME_SOBRENOME - ATENDIMENTOS - DD-MM-AA.xlsm`

### **"Excel está aberto" sempre aparece**

- Use a Opção 2 para trabalhar com arquivo temporário
- Ou feche todas as instâncias do Excel

### **"Arquivo já está sendo usado"**

- Feche a planilha no Excel
- Ou use a Opção 2 para criar temporário

### **Programa não abre**

- Execute como Administrador (botão direito → Executar como administrador)
- Verifique se não está bloqueado pelo Windows (Propriedades → Desbloquear)

---

## 💡 Dicas de Uso

### **Para melhor desempenho:**

- Feche o Excel antes de executar (Opção 1)
- Mantenha arquivos organizados na pasta `filhas/`
- Verifique os logs em caso de erro

### **Para trabalhar simultaneamente:**

- Use Opção 2 (arquivo temporário)
- Consulte o temporário enquanto trabalha
- Execute novamente após fechar Excel principal

### **Para backup:**

- Salve a pasta `backup/` periodicamente
- Mantenha cópias dos arquivos originais
- Verifique o log após cada execução

---

## 📞 Suporte

Em caso de dúvidas ou problemas:

1. Verifique o arquivo `log_compilacao.txt`
2. Certifique-se que os arquivos seguem o padrão de nomenclatura
3. Teste com um arquivo de exemplo primeiro

---

## 📄 Informações Técnicas

- **Versão:** 2.0
- **Plataforma:** Windows 64-bit
- **Compilado com:** PyInstaller 6.16.0
- **Bibliotecas:** pandas, openpyxl
- **Formato de saída:** Excel (.xlsx)

---

## ✅ Checklist de Distribuição

Antes de distribuir, verifique:

- [ ] `atualizar_planilhas.exe` está presente
- [ ] Pasta `filhas/` existe (mesmo que vazia)
- [ ] Instruções de nomenclatura estão claras
- [ ] Testado em outro PC/pasta
- [ ] README.md incluído (este arquivo)

---

**Desenvolvido para automação de consolidação de planilhas de atendimentos**  
_Versão Portátil - Funciona em qualquer pasta, qualquer PC Windows_
