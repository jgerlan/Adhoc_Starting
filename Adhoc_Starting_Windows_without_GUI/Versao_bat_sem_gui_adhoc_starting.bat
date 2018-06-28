@echo off
echo Iniciando modo adhoc...
echo nome da lan: redeadhoc
netsh wlan set hostednetwork mode=allow ssid=redeadhoc key=12345678
echo Rede iniciada!
netsh wlan start hostednetwork
echo Deseja cancelar conexao? pressione enter para isso!
pause
netsh wlan stop hostednetwork