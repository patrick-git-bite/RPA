#!/bin/bash

# Este script automatiza o processo de upload do seu projeto para o GitHub.
# Certifique-se de que o Git está instalado e configurado em sua máquina.

# URL do seu repositório GitHub
REPO_URL="https://github.com/patrick-git-bite/RPA"

# Mensagem do primeiro commit
COMMIT_MESSAGE="Primeiro commit: Projeto RPA de Baixa de AFs"

# 1. Inicializar o repositório Git localmente
echo "Inicializando o repositório Git..."
git init

# 2. Adicionar todos os arquivos do projeto ao stage
echo "Adicionando arquivos ao stage..."
git add .

# 3. Realizar o primeiro commit
echo "Realizando o primeiro commit..."
git commit -m "$COMMIT_MESSAGE"

# 4. Adicionar o repositório remoto do GitHub
echo "Adicionando o repositório remoto..."
git remote add origin $REPO_URL

# 5. Renomear a branch principal para 'main'
echo "Renomeando a branch para 'main'..."
git branch -M main

# 6. Enviar os arquivos para o GitHub
echo "Enviando arquivos para o GitHub..."
git push -u origin main

echo "\nProcesso de upload para o GitHub concluído!"
echo "Você pode precisar inserir suas credenciais do GitHub (nome de usuário e Personal Access Token - PAT) na primeira vez."


