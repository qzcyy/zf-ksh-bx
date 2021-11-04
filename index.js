const fs = require('fs')
const baiduSdk = require('./baidu-sdk')
const Excel = require('exceljs');
const moment = require('moment')
const path = require('path')
let {createReport} = require('docx-templates');

let api_key = '';
let secret_key = ''
let FILE_PATH = './data'
const DOCX_IMAGE_SIZE = 4
const EMPTY_BASE64 = fs.readFileSync(path.join(__dirname, 'template/empty.jpg')).toString('base64')
let resultDataTemp = {};
let resultData = [];
let resultImage = []
let cantOCRImg = []
let totalPrice = 0

function ckeckPayType(words_result) {
    if (words_result.find(function (item) {
        return item.words === 'AA收款'||item.words === '商家订单号'
    })) {
        return 'zfb'
    } else if (words_result.find(function (item) {
        return item.words === '发起群收款'
    })) {
        return 'wechart'
    }else if (words_result.find(function (item) {
        return item.words === '美团支付'
    })) {
        return 'mt'
    }
}

//数组二维化
function chunk(item, size) {
    if (item.length <= 0 || size <= 0) {
        return item;
    }
    let chunks = [];
    for (let i = 0; i < item.length; i = i + size) {
        chunks.push(item.slice(i, i + size));
    }
    return chunks
}


async function parseDataAfter(parseData) {
    //只报销30
    parseData.price=parseData.price>30?30:parseData.price
    parseData.showTime = parseData.time.slice(0, 10)
    resultDataTemp[parseData.showTime] = resultDataTemp[parseData.showTime] || []
    if (!resultDataTemp[parseData.showTime].find(function (item) {
        return item.time === parseData.time && item.price === parseData.price
    })) {
        totalPrice += parseData.price
        resultDataTemp[parseData.showTime].push(parseData)
    }

}

(async () => {



})()

module.exports=async (cmdPath,keyCache)=>{
    api_key=keyCache.api_key
    secret_key=keyCache.secret_key
    const workbook = new Excel.Workbook();
    //清空未识别图片文件夹
    if(fs.existsSync('未识别图片')){
        const oldFile=await fs.readdirSync('未识别图片')
        for(const oldFileItem of oldFile){
            fs.unlinkSync('未识别图片/'+oldFileItem)
        }
    }

    await workbook.xlsx.readFile(path.join(__dirname+'/template/月加班餐费明细表.xlsx'));
    const worksheet = workbook.getWorksheet('招待');


    let {data: authData} = await baiduSdk.getAuth(api_key, secret_key)
    let {access_token} = authData
    if(cmdPath) FILE_PATH=cmdPath
    let dirs = fs.readdirSync(FILE_PATH)
    for (const dir of dirs.filter(dir=>fs.statSync(dir).isDirectory())) {
        let dirsFile = fs.readdirSync(FILE_PATH + '/' + dir)
        //过滤打车
        for (const file of dirsFile.filter(file=>file.indexOf('车')===-1)) {
            let img = fs.readFileSync(`${FILE_PATH}/${dir}/${file}`)
            console.log(`查询${file}`);
            let {data: body} = await baiduSdk.getImageDetail(access_token, {
                image: img.toString('base64')
            })
            let words_result = body.words_result
            if (ckeckPayType(words_result) === 'zfb') {
                //支付宝
                let parseData = {
                    price: Math.abs(words_result[words_result.findIndex(function (item) {
                        return item.words == '付款方式'
                    }) - 2].words),
                    time: words_result[words_result.findIndex(function (item) {
                        return item.words == '创建时间'
                    }) + 1].words.replace('年','-').replace('月','-').replace('日',''),
                    // orderNo: words_result[words_result.findIndex(function (item) {
                    //     return item.words == '订单号'
                    // }) + 1].words,
                    type: 'zfb',
                    local: file,
                    owner: dir
                }
                await parseDataAfter(parseData)

            } else if (ckeckPayType(words_result) === 'wechart') {
                //微信
                let parseData = {
                    price: Math.abs(words_result[words_result.findIndex(function (item) {
                        return item.words == '当前状态'
                    }) - 1].words),
                    time: words_result[words_result.findIndex(function (item) {
                        return item.words == '支付时间'
                    }) + 1].words.replace('年','-').replace('月','-').replace('日',''),
                    // orderNo: words_result[words_result.findIndex(function (item) {
                    //     return item.words == '订单号'
                    // }) + 1].words,
                    type: 'wechart',
                    local: file,
                    owner: dir
                }
                await parseDataAfter(parseData)
            }else {
                cantOCRImg.push(`${FILE_PATH}/${dir}/${file}`)
            }
        }
    }
    let rows = []
    const _resultDataTempCache = Object.keys(resultDataTemp).sort((a, b) => moment(a) - moment(b))
    for (const key of _resultDataTempCache) {
        const index = _resultDataTempCache.indexOf(key);
        resultImage = resultImage.concat(resultDataTemp[key].map(tempItem => `${FILE_PATH}/${tempItem.owner}/${tempItem.local}`))
        rows.push([
            index + 1,
            key,
            resultDataTemp[key].reduce((acc, cur) => acc + cur.owner + ',', ''),
            '济南',
            resultDataTemp[key].length,
            resultDataTemp[key].reduce((acc, cur) => Number(acc) + Number(cur.price), ''),
            ''
        ])
        resultData.push({
            time: key,
            value: resultDataTemp[key]
        })

    }
    worksheet.getCell('A15').value = `上述费用，共____${rows.length}____笔，金额共计___${totalPrice}___元。`
    await worksheet.insertRows(7, rows, 'o+')
    //把站位字体删除
    // await worksheet.spliceRows(7+rows.length, 10);
    await workbook.xlsx.writeFile(`自动生成-${moment().format('YYYY-MM')}加班餐费明细表.xlsx`);
    console.log('excel文档生成成功');

    console.log('执行插入base64编码图片' + resultImage);
    let docxTemplate = await fs.readFileSync(path.join(__dirname+'/template/wordTemp.docx'))
    let docxData = []
    let _docxData = []
    for (let resultImageItem of resultImage) {
        let _tem = await fs.readFileSync(resultImageItem).toString('base64')
        docxData.push(_tem)
        _docxData.push(resultImageItem)
    }
    let emptySize = DOCX_IMAGE_SIZE - docxData.length % DOCX_IMAGE_SIZE
    for (let i = 0; i < emptySize; i++) {
        docxData.push(EMPTY_BASE64)
    }
    docxData = chunk(docxData, DOCX_IMAGE_SIZE)
    const docxBuffer = await createReport({
        template: docxTemplate,
        // output: './test.docx',
        data: {
            userImg: docxData
        },
        additionalJsContext: {
            showImg: (base64Arr, index) => {
                return {width: 6, height: 12, data: base64Arr[index], extension: '.jpg'};
            },
        }
    });
    fs.writeFileSync(`自动生成-${moment().format('YYYY-MM')}月加班餐费用粘贴单.docx`, docxBuffer)
    console.log('word文档生成成功');

    if(cantOCRImg.length){
        if(!fs.existsSync('未识别图片')){
            await fs.mkdirSync('未识别图片')
        }
        for(const cantOCRImgItem of cantOCRImg){
            fs.copyFileSync(cantOCRImgItem,'未识别图片/'+cantOCRImgItem.substring(cantOCRImgItem.lastIndexOf('/'),cantOCRImgItem.length))
        }
        console.log(cantOCRImg);
        console.error('存在未识别图片，请人工分配')
    }
}
