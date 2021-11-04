#!/usr/bin/env node

process.title = 'c';
var program = require('commander');
var path = require('path');
var createXlsxAndWord = require('../index')
var fs=require('fs')
program
    .version(require('../package').version,'-v, --version')
    .option('-s, --setKey <keys...>','setKey')
    .option('-c, --create','create Xlsx and Word')
    .action(function (cmd) {
        if(cmd.setKey){
            fs.writeFileSync(path.join(__dirname+'/keyCache.ini'),JSON.stringify({
                api_key:cmd.setKey[0],
                secret_key:cmd.setKey[1]
            }),"utf-8")
        }else if(cmd.create){
            var keyCache=fs.readFileSync(path.join(__dirname+'/keyCache.ini'),"utf-8")
            if(!keyCache){
                console.log('未查询到有效apiKEY，请到百度开房平台申请图片识别KEY')
                process.exit(1);
            }
            createXlsxAndWord(process.cwd(),JSON.parse(keyCache))
        }
    })
    .parse(process.argv);
