import { MongoClient } from 'mongodb';
import { exportExcel } from '@cnopendata/formatexcel';
import path from 'path';

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
async function mongodbtoExcel(uri, dbName, colName, headerKeys, dirPath, filename, func = (doc) => doc, concurrency = 8, rowCount = 50000, skip = 0) {
    const client = new MongoClient(uri);
    await client.connect();
    const col = client.db(dbName).collection(colName);
    let docs = [];
    skip = skip - (skip % concurrency);
    let count = 0 + (skip / concurrency);
    await col.find().skip(skip * rowCount).forEach(async function (doc) {
        if (docs.length >= rowCount * concurrency) {
            Promise.all([...Array(concurrency).keys()].map((i) => exportExcel(headerKeys, docs.slice(i * rowCount, (i + 1) * rowCount), path.join(dirPath, `${filename}-${count * concurrency + i}.xlsx`)).then(console.log(`成功写入${filename}-${count * concurrency + i}.xlsx`))))
            docs = docs.slice(concurrency*rowCount);
            count += 1;
        }
        doc = func(doc);
        if (doc) {
            if (typeof (doc?.[Symbol.iterator]) === 'function') {
                docs.push(...doc);
            } else {
                docs.push(doc);
            }
        }
    });
    await Promise.all([...Array(Math.ceil(docs.length / rowCount)).keys()].map((i) => exportExcel(headerKeys, docs.slice(i * rowCount, (i + 1) * rowCount), path.join(dirPath, `${filename}-${count * concurrency + i}.xlsx`))));
    await client.close();
}

export { mongodbtoExcel }