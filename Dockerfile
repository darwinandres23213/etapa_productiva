# Consulta Etapa Productiva SENA — imagen lista para Render
FROM nginx:1.27-alpine

# Render inyecta PORT; el entrypoint oficial de nginx procesa templates/
ENV PORT=8080

# Quitar default y montar config con puerto dinámico
RUN rm -f /etc/nginx/conf.d/default.conf
COPY nginx/default.conf.template /etc/nginx/templates/default.conf.template

# App estática
COPY index.html styles.css app.js excel-import.js sena-logo.svg data.json /usr/share/nginx/html/

# Permisos de lectura
RUN chmod -R a+r /usr/share/nginx/html

EXPOSE 8080

# El entrypoint de nginx:alpine ya hace envsubst de /etc/nginx/templates
CMD ["nginx", "-g", "daemon off;"]
