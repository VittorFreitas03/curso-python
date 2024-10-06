import pyautogui
import time
import os
import webbrowser  # Importando o módulo webbrowser

# Caminho para a sua planilha Excel
file_path = r'C:\teste\PLANILHACALCULODEDISTANCIAS.xlsm'  # Altere para o seu caminho

# Abrir a planilha do Excel
os.startfile(file_path)

# Aguardar alguns segundos para o Excel abrir
time.sleep(5)  # Ajuste o tempo se necessário

# Copiar o conteúdo da célula A1
pyautogui.hotkey('ctrl', 'g')  # Abre a caixa de diálogo "Ir para"
time.sleep(1)  # Esperar a caixa de diálogo abrir
pyautogui.write('A1')  # Digita A1
pyautogui.press('enter')  # Vai para a célula A1
time.sleep(1)  # Esperar a célula ser selecionada
pyautogui.hotkey('ctrl', 'c')  # Copia o conteúdo da célula A1

# Copiar o conteúdo da célula B1
pyautogui.hotkey('ctrl', 'g')  # Abre a caixa de diálogo "Ir para" novamente
time.sleep(1)  # Esperar a caixa de diálogo abrir
pyautogui.write('B1')  # Digita B1
pyautogui.press('enter')  # Vai para a célula B1
time.sleep(1)  # Esperar a célula ser selecionada
pyautogui.hotkey('ctrl', 'c')  # Copia o conteúdo da célula B1

# Abrir o navegador e acessar um site específico
link_do_site = 'https://rotasbrasil.com.br'  # Altere para o seu link
webbrowser.open(link_do_site)

# Aguardar alguns segundos para o site abrir
time.sleep(10)  # Ajuste o tempo se necessário


pyautogui.click(x=114, y=207)

# Pressionar Windows + V para abrir o histórico da área de transferência
pyautogui.hotkey('win', 'v')

# Aguardar um tempo para que o histórico seja aberto
time.sleep(1)  # Ajuste o tempo se necessário

# Pressionar a seta para baixo para selecionar o item mais recente
pyautogui.press('down')

# Aguardar um momento antes de pressionar Enter
time.sleep(0.5)

# Pressionar Enter para colar o item selecionado
pyautogui.press('enter')

time.sleep (1)

pyautogui.press('down')

pyautogui.press('enter')

time.sleep (4)

pyautogui.click(x=46, y=252)

# Pressionar Windows + V para abrir o histórico da área de transferência
pyautogui.hotkey('win', 'v')

# Aguardar um tempo para que o histórico seja aberto
time.sleep(1)  # Ajuste o tempo se necessário


# Aguardar um momento antes de pressionar Enter
time.sleep(0.5)

pyautogui.press('enter')

time.sleep(1)

pyautogui.press('down')

pyautogui.press('enter')

time.sleep(6)


pyautogui.click(x=159, y=421)
