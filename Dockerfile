## build

FROM node:18-alpine as build
WORKDIR /usr/src/app
COPY . ./
RUN npm install && npm run build

FROM nginx
RUN rm -rf /usr/share/nginx/html/*
COPY --from=build /usr/src/app/dist /usr/share/nginx/html
