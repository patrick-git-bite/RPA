import win32com.client
import re
import pandas as pd
import os
import subprocess

def processar_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Pasta de entrada
    mensagens = inbox.Items
    mensagens.Sort("[ReceivedTime]", True)
    mensagens = mensagens.Restrict("[UnRead] = true")  # Somente não lidos

    lista_afs = []

    for mail in list(mensagens):
        try:
            assunto = mail.Subject or ""
            
            # Expressão regular para verificar o assunto
            if re.search(r"baixa[s]?(?:\s+de)?\s+af", assunto, re.IGNORECASE):
                corpo = mail.Body or ""

                blocos = re.findall(
                    r"(Restaurante[:\- ]+(\d{4})(.*?)?)(?=Restaurante[:\- ]+\d{4}|$)",
                    corpo,
                    re.IGNORECASE | re.DOTALL
                )

                if not blocos:
                    print("⚠️ Nenhum restaurante encontrado no corpo do e-mail.")
                    mail.UnRead = False
                    continue

                print(f"\n📧 E-mail: {assunto}")

                for bloco_completo, restaurante, trecho_afs in blocos:
                    afs = re.findall(r"^\s*(\d{5,7})\s*$", trecho_afs, re.MULTILINE)
                    if afs:
                        print(f"📍 Restaurante: {restaurante}")
                        print(f"🧾 AFs encontradas: {afs}")
                        for af in afs:
                            lista_afs.append({"Restaurante": restaurante, "AF": af, "Status": "Concluído"})
                    else:
                        print(f"⚠️ Nenhuma AF encontrada para o restaurante {restaurante}.")

                mail.UnRead = False
                print("📩 E-mail marcado como lido.")

                try:
                    resposta = mail.Reply()
                    resposta.Body = (
                        "Olá,\n\n"
                        "A baixa das AFs foi registrada com sucesso. "
                        "Por favor, confira os registros no sistema. "
                        "Se houver divergência, reenvie no mesmo formato.\n\n"
                        "Obrigado!"
                    )
                    resposta.Send()
                    print("✉️ E-mail de confirmação enviado.")
                except Exception as e:
                    print(f"⚠️ Não foi possível enviar a resposta: {e}")
        except Exception as e:
            print(f"❌ Erro ao processar e-mail: {e}")

    return lista_afs

def atualizar_historico(lista_afs):
    caminho_historico = os.path.join(os.getenv("USERPROFILE"), "Desktop", "testes", "historico_af.xlsx")
    df_novo = pd.DataFrame(lista_afs)

    if not df_novo.empty:
        if os.path.exists(caminho_historico):
            try:
                df_hist = pd.read_excel(caminho_historico, sheet_name="Histórico", dtype={"Restaurante": str, "AF": str})
                df_hist_final = pd.concat([df_hist, df_novo], ignore_index=True)
                df_hist_final.drop_duplicates(subset=["Restaurante", "AF"], inplace=True)
            except Exception as e:
                print(f"⚠️ Erro ao ler histórico existente, criando novo: {e}")
                df_hist_final = df_novo
        else:
            df_hist_final = df_novo

        with pd.ExcelWriter(caminho_historico, engine='openpyxl', mode='w') as writer:
            df_hist_final.to_excel(writer, index=False, sheet_name="Histórico")
        print("🗃️ Histórico de AFs atualizado.")
    else:
        print("⚠️ Não há dados para atualizar o histórico.")

def salvar_dados_af(lista_afs):
    caminho_dados = os.path.join(os.getenv("USERPROFILE"), "Desktop", "testes", "dados_af.xlsx")
    df_novo = pd.DataFrame(lista_afs)

    if not df_novo.empty:
        with pd.ExcelWriter(caminho_dados, engine='openpyxl', mode='w') as writer:
            df_novo.to_excel(writer, index=False, sheet_name="AFs Recebidas")
        print("📄 dados_af.xlsx sobrescrito com AFs recentes.")
    else:
        print("⚠️ Não há dados para salvar no arquivo de AFs.")

def executar_uipath():
    caminho_bat = os.path.join(os.getcwd(), "executar_uipath.bat")
    if os.path.exists(caminho_bat):
        try:
            subprocess.Popen(caminho_bat, shell=True)
            print("🚀 UiPath iniciado automaticamente.")
        except Exception as e:
            print(f"⚠️ Erro ao iniciar UiPath: {e}")
    else:
        print(f"⚠️ Arquivo {caminho_bat} não encontrado. UiPath não foi iniciado.")

if __name__ == "__main__":
    print("🔍 Iniciando verificação de e-mails não lidos com 'baixa de af' no assunto...")
    lista_afs = processar_emails()

    # Verificar se encontrou algum e-mail
    if lista_afs:
        print(f"✅ {len(lista_afs)} AFs encontradas. Atualizando histórico e salvando dados...")
        atualizar_historico(lista_afs)
        salvar_dados_af(lista_afs)
        executar_uipath()
    else:
        print("ℹ️ Nenhum e-mail novo para processar.")
    
    print("✅ Processo finalizado.")
