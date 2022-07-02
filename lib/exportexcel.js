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
 * @param {(Object)=> Object} [func=(doc)=>doc] - 处理document并返回 键要与headerKeys的key对应, 默认直接返回
 * @param {number} [concurrency=8] - 线程数, 默认8
 * @param {number} [rowCount=50000] - excel的最大行数, 默认50000
 */
async function mongodbtoExcel(uri, dbName, colName, headerKeys, dirPath, filename, func = (doc) => doc, concurrency = 8, rowCount = 50000) {
    const client = new MongoClient(uri);
    await client.connect();
    const col = client.db(dbName).collection(colName);
    let docs = [];
    let count = 0;
    await col.find().forEach(async function (doc) {
        if (docs.length === rowCount * concurrency) {
            Promise.all([...Array(concurrency).keys()].map((i) => exportExcel(headerKeys, docs.slice(i*rowCount, (i+1)*rowCount), path.join(dirPath, `${filename}-${count*concurrency + i}.xlsx`)).then(console.log(`成功写入${filename}-${count*concurrency + i}.xlsx`))))
            docs = [];
            count += 1;
        }
        docs.push(func(doc));
    });
    await Promise.all([...Array(Math.ceil(docs.length / rowCount)).keys()].map((i) => exportExcel(headerKeys, docs.slice(i * rowCount, (i + 1) * rowCount), path.join(dirPath, `${filename}-${count * concurrency + i}.xlsx`))));
    await client.close();
}

export { mongodbtoExcel }