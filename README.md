# 📊 Gestor de Vendas - Exportar para Excel

Um aplicativo desktop completo para gerenciar vendas e exportar dados automaticamente para planilhas Excel, com interface gráfica intuitiva.

## ✨ Funcionalidades

- ✅ **Adicionar vendas** com produto, quantidade e valor unitário
- ✅ **Cálculo automático** de totais por venda
- ✅ **Visualização em tempo real** de todas as vendas em tabela
- ✅ **Remover vendas** individuais ou limpar tudo
- ✅ **Exportar para Excel** com formatação profissional
- ✅ **Total automático** da soma de todas as vendas
- ✅ **Interface gráfica amigável** (sem linhas de comando)

## 🚀 Como Usar

### Opção 1: Executável (Recomendado)

1. Acesse a pasta `dist/`
2. Clique duas vezes em `Gestor_Vendas.exe`
3. Preencha os dados e clique em "Adicionar"
4. Clique em "📊 Exportar para Excel" para salvar

**Vantagens:**
- Não requer Python instalado
- Funciona em qualquer Windows
- ~11 MB de tamanho

### Opção 2: Script Python

```bash
# Clonar o repositório
git clone https://github.com/seu-usuario/gestor-vendas.git
cd gestor-vendas

# Criar ambiente virtual
python -m venv venv
.\venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/Mac

# Instalar dependências
pip install -r requirements.txt

# Executar
python py001.py
```

## 📋 Requisitos

**Para o executável:**
- Windows 7 ou superior
- ~20 MB de espaço livre

**Para rodar o script:**
- Python 3.10+
- openpyxl (instalado automaticamente)

## 📁 Estrutura do Projeto

```
gestor-vendas/
├── py001.py              # Script principal
├── requirements.txt      # Dependências Python
├── README.md            # Este arquivo
├── .gitignore           # Arquivos a ignorar no Git
└── dist/
    ├── Gestor_Vendas.exe    # Executável compilado
    └── README.txt           # Instruções de uso
```

## 💾 Dependências

```
openpyxl>=3.0.0
```

Todas as dependências estão listadas em `requirements.txt`

## 📊 Formato do Excel Exportado

O arquivo Excel gerado possui:

- **Cabeçalho formatado** (azul com texto branco)
- **Colunas**: ID | Produto | Quantidade | Valor Unitário | Total | Data
- **Linha de total** (destacada em amarelo)
- **Largura automática** das colunas
- **Data e hora** do registro

## 🔧 Desenvolvimento

### Instalação para Desenvolvimento

```bash
# Clonar
git clone https://github.com/seu-usuario/gestor-vendas.git
cd gestor-vendas

# Criar venv
python -m venv venv
.\venv\Scripts\activate

# Instalar dependências
pip install -r requirements.txt
```

### Compilar Executável

```bash
pip install pyinstaller

pyinstaller --onefile --windowed --name "Gestor_Vendas" py001.py
```

O executável será gerado em `dist/Gestor_Vendas.exe`

## 📝 Como Contribuir

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 🎯 Roadmap de Melhorias

- [ ] Importar dados de arquivo Excel
- [ ] Gráficos de vendas
- [ ] Base de dados SQLite
- [ ] Editar vendas registradas
- [ ] Filtros por data
- [ ] Relatórios customizáveis
- [ ] Modo escuro

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 👤 Autor

**Seu Nome**
- GitHub: [@seu-usuario](https://github.com/seu-usuario)
- Email: seu.email@exemplo.com

## 🤝 Suporte

Encontrou um bug? Abra uma [issue](https://github.com/seu-usuario/gestor-vendas/issues)!

## 📞 Contato

Dúvidas? Perguntas? Sugestões?

- 📧 Email: seu.email@exemplo.com
- 💬 GitHub Discussions

---

**Versão:** 1.0  
**Última atualização:** 23/03/2026
