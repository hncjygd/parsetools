## 介绍

该模块个人使用，主要用于解析数据并保存到数据库中，以及从数据库中提取数据到excel中。


## API介绍

```javascript
/**
 * 将mongodb数据导出到excel中。
 * @param {string} uri Mongodb Server URI
 * @param {string} dbName dbName
 * @param {string} colName colName
 * @param {{header:string,key:string}[]} headerKeys - 数组，元素为{header: 表头, key: 列键}
 * @param {string} dirPath - 要写入的文件夹路径 
 * @param {string} filename - 文件名前缀，后自动添加0 1 2.. 序列 
 * @param {(Object)=> Object} [func=(doc)=>doc] - 处理document并返回 键要与headerKeys的key对应, 默认直接返回
 * @param {number} [concurrency=8] - 线程数, 默认8
 * @param {number} [rowCount=50000] - excel的最大行数, 默认50000
 */
async function mongodbtoExcel(uri, dbName, colName, headerKeys, dirPath, filename, func = (doc) => doc, concurrency = 8, rowCount = 50000)
```