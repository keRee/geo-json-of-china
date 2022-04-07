const request = require('request')
const fs = require('fs')
const path = require('path')
const Excel = require('exceljs')
const Chalk = require('chalk')

const GeoJSONTarget = path.resolve(__dirname, '../dist/')
const AdCodeExcel = path.resolve(__dirname, './assets/AMap_adcode_citycode.xlsx')
const AdCodeExcel2Json = path.resolve(__dirname, './assets/AMap_adcode_citycode.json')

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
        App.tryGetGeoJSON(j, len, data)
    },
    tryGetGeoJSON(j, len, data) {
        App.getGeoJSON(data[j], () => {
            j++
            const rate = j === len ? 100 : ((j / len) * 100).toFixed(2)
            console.log(Chalk.green(`Tips: 第 ${j} 条数据写入完毕, 已完成 `), `${Chalk.red(rate)} %`)
            j < len && App.tryGetGeoJSON(j, len, data)
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
            fs.writeFileSync(`${GeoJSONTarget}/${fileName}`, body)
            typeof cb === 'function' && cb()
        })
    }
}
App.init()