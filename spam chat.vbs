set shell = createobject ("wscript.shell")

strtext = inputbox ("Digite abaixo a mensagem que gostaria de spammar")
strtimes = inputbox ("Quantas vezes deseja spammar?")
strspeed = inputbox ("O quao rapido deseja spammar? (1000 = 1 por segundo, 100 = 10 por segundo etc. Recomendado 1000.)")
strtimeneed = inputbox ("Quanto tempo precisa para selecionar o chat que deseja spammar?")
If not isnumeric (strtimes & strspeed & strtimeneed) then
msgbox "Você digitou algo diferente de numeros/vezes, velocidade e/ou tempo necessario. Fechando o Programa."
wscript.quit
End If
strtimeneed2 = strtimeneed * 1000
do
msgbox "Voce tem " & strtimeneed & " segundos para ir aonde deseja iniciar o spam."
wscript.sleep strtimeneed2
shell.sendkeys ("Spambot ativado" & "{enter}")
for i=0 to strtimes
shell.sendkeys (strtext & "{enter}")
wscript.sleep strspeed
Next
shell.sendkeys ("Spambot desativado" & "{enter}")
wscript.sleep strspeed * strtimes / 10
returnvalue=MsgBox ("Deseja spammar com as mesmas informacoes novamente?",36)
If returnvalue=6 Then
Msgbox "Ok o Bot ira iniciar novamente!"
End If
If returnvalue=7 Then
msgbox "Spambox esta desligando!"
wscript.quit
End IF
loop