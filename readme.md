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
 * @param {(Object)=> Object|Object[]} [func=(doc)=>doc] - 处理document并返回文档本身，返回值为false的将被抛弃。键要与headerKeys的key对应, 默认直接返回。如果返回一个可迭代对象会将所有元素添加进去
 * @param {number} [concurrency=8] - 线程数, 默认8
 * @param {number} [rowCount=50000] - excel的最大行数, 默认50000
 * @param {number} [skip=0] - 跳过解析的数量，主要用于当导出出错时重新指定导出的起点位置，位置为skip*rowCount, 所以填写需要重新开始导出的文件的序号即可，其他参数需要保持不变, 由于异步，所以只能重concurrency的倍数处重新提取可能会覆盖部分文件
 */
async function mongodbtoExcel(uri, dbName, colName, headerKeys, dirPath, filename, func = (doc) => doc, concurrency = 8, rowCount = 50000, skip = 0)
```