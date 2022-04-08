const request = require('request')
const fs = require('fs')
const path = require('path')
const Excel = require('exceljs')
const Chalk = require('chalk')

const GeoJSONTarget = path.resolve(__dirname, '../dist/')
const AdCodeExcel = path.resolve(__dirname, './assets/AMap_adcode_citycode.xlsx')
const AdCodeExcel2Json = path.resolve(__dirname, './assets/AMap_adcode_citycode.json')
const InvalidDataJson = path.resolve(__dirname, './assets/invalidData.json')
const CityDataJson = path.resolve(__dirname, './assets/cityData.json')
const ErrorBodyJson = path.resolve(__dirname, './assets/errorBody.json')

const invalidData = []
const cityData = []
let errorBody = ''
const App = {
    baseUrl: 'https://geo.datav.aliyun.com/areas_v3/bound',
    init() {
        this.getAllAdCode()
    },
    getAllAdCode: async () => {
        const Workbook = new Excel.Workbook()
        const workbook = await Workbook.xlsx.readFile(AdCodeExcel)
        // 获取第1个sheet
        const worksheet = workbook.getWorksheet(1)
        let i = 2;
        let data = []
        let cur = ''
        do {
            const row = worksheet.getRow(i++)
            const d = {
                name: row.getCell(1).value || '',
                adCode: row.getCell(2).value || '',
                cityCode: row.getCell(3).value || ''
            }
             if (d.name && d.adCode) {
                data.push(d)
                cur = d.name
             } else {
                cur = ''
             }
        } while(cur)
        console.log(Chalk.green(`Tips: ${data.length} 条数据读取完毕！`))
        fs.writeFileSync(AdCodeExcel2Json, JSON.stringify(data, null, '\t'), {
            encoding: 'utf8'
        })
        console.log(Chalk.green(`Tips: ${data.length} 条数据写入完毕！`))
        
        let j = 0
        let len = data.length
        // let len = 1000
        App.tryGetGeoJSON(j, len, data, 0)
    },
    tryGetGeoJSON(j, len, data, trueIndex) {
        App.getGeoJSON(data[j], (hasGeoJSON) => {
            j++
            hasGeoJSON && trueIndex++
            const rate = j === len ? 100 : ((j / len) * 100).toFixed(2)
            console.log(Chalk.green(`Tips: 第 ${trueIndex} 条数据写入完毕，${j - trueIndex}条无效，已完成 `), `${Chalk.red(rate)} %`)
            fs.writeFileSync(`${InvalidDataJson}`, JSON.stringify(invalidData, null, '\t'))
            fs.writeFileSync(`${CityDataJson}`, JSON.stringify(cityData, null, '\t'))
            fs.writeFileSync(`${ErrorBodyJson}`, errorBody)
            if (j < len) {
                App.tryGetGeoJSON(j, len, data, trueIndex)
            }
        })
    },
    getGeoJSON({ adCode, name }, cb) {
        const self = this
        const fileName = `${adCode}.json`
        request(`${self.baseUrl}/${fileName}`, function (error, response, body) {
            if (error) {
                console.log(Chalk.red(error))
                return
            }
            if (!fs.existsSync(GeoJSONTarget)) {
                fs.mkdirSync(GeoJSONTarget)
            }
            const hasGeoJSON = /^\{/.test(body)
            if (hasGeoJSON) {
                cityData.push({
                    k: adCode,
                    v: name
                })
                fs.writeFileSync(`${GeoJSONTarget}/${fileName}`, body)
                fs.writeFileSync(`${GeoJSONTarget}/${name}.json`, body)
            } else {
                invalidData.push({
                    name,
                    adCode
                })
                errorBody = body
            }
            const delay = 1 + Math.ceil(Math.random() * 5);
            setTimeout(() => {
                typeof cb === 'function' && cb(hasGeoJSON)
            }, delay * 1000)
        })
    }
}
App.init()