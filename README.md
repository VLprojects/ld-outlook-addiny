Необходимо: 

1 - создать на вашем хостинге livedigital.space поддомен (например ms.livedigital.space) и прописать его в 
webpack.config.js
const urlProd = "https://ld-outlook-addin.onrender.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

2 - Использовать для прода команды
nmp install 
nmp run build

3 - загрузить на хостинг (в корень поддомена) папку dist
