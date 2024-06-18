FROM alpine:latest 

ENV PLUGIN_DIR=/var/www/onlyoffice/documentserver/sdkjs-plugins/biyue

RUN pwd
WORKDIR /biyue
COPY . .

ENTRYPOINT ["sh"]
CMD ["-c", "mkdir -p ${PLUGIN_DIR} && rm -rf {PLUGIN_DIR}/* && cp -r /biyue/* ${PLUGIN_DIR} && tail -f /dev/null"]
