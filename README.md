Необходимо:

1. Создать на вашем хостинге livedigital.space поддомен (например ms.livedigital.space)

2. Добавить разрешения cors для созданного поддомена в сервисе АПИ https://moodhood-api.livedigital.space/v1/

3. Прописать поддомен в webpack.config.js

   const urlProd = "https:// YOURSUBDOMAIN.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

4. Использовать для прода команды ( версия node.js min 18)

   nmp install

   nmp run build

5. Загрузить в корень поддомена содержимое папки dist
