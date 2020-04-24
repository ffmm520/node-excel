// 导出excel数据 demo
const XLSX = require('xlsx')
const stream = require('stream')
const express = require('express')
const app = express()

app.get('/', (req, res) => {
	res.send('index page')
})
// 访问localhost:3002/xlsx 即可下载
app.get('/xlsx', (req, res) => {
	// 1. 二维数组，用作数据源
	const ws_data = [
		['id', 'name', 'role', 'path', 'state'],
		[1, '首页', 'user', '/', 1],
		[2, '全部文章', 'user', '/article', 1],
	]

	// 2. 使用数据源创建sheet
	const worksheet = XLSX.utils.aoa_to_sheet(ws_data)
	// 3. 创建空的表格簿
	let workbook = XLSX.utils.book_new()

	// 4. 将刚创建的 sheet 追加到表格簿中，并为这个新创建的 sheet 命名为 SheetJS
	XLSX.utils.book_append_sheet(workbook, worksheet, 'SheetJS')

	// 5. 执行写入方法将内存中的表格簿写入到文件
	// XLSX.writeFile(workbook, 'out.xlsx')

	const fileContents = XLSX.write(workbook, {
		type: 'buffer',
		bookType: 'xlsx',
		bookSST: false,
	})

	var readStream = new stream.PassThrough()
	readStream.end(fileContents)

	let fileName = 'text.xlsx'
	res.set('Content-disposition', 'attachment; filename=' + fileName)
	res.set('Content-Type', 'text/plain')

	readStream.pipe(res)
})

app.listen(3002, () => {
	console.log('start on port 3002')
})
