

/*  */



import odbc from 'odbc';
import path  from 'path';
// //=== กรณีไฟล์ไม่อยู่ในโฟลเดอร์โปรเจกต์ - Absolute path
// const dbFilePath = "C:/Users/wasankds/Documents/Database1.accdb";

//=== กรณีไฟล์อยู่ในโฟลเดอร์โปรเจกต์ - Relative path
const rootDir = path.resolve();

// accDb
const dbFilePath = path.join(rootDir, 'files', 'WK_MDB.mdb');

// mdb
// const dbFilePath = path.join(rootDir, 'files', 'WK_ACCDB.accdb');

//=== 
const msDriver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
const connectionString = `Driver=${msDriver};DBQ=${dbFilePath};`;
async function readAccdb() {
  let connection = await odbc.connect(connectionString);
  try {
    const result = await connection.query('SELECT * FROM member');
    // console.log('result ===> ' , result);
    const rows = result.filter(item => typeof item.ID !== 'undefined');
    console.log('rows ===> ' , rows);


    // const result = await connection.query('SELECT * FROM qryrptEmployeeEmailList');
    // console.log('result ===> ', result);

  } catch (error) {
    console.error('เกิดข้อผิดพลาดในการอ่านฐานข้อมูล:', error);
  } finally {
    if (connection) {
      await connection.close();
    }
  }
}

readAccdb();




// import MDBReader from 'mdb-reader';
// import fs from 'fs';

// const dbFilePath = "C:\\Users\\wasankds\\Documents\\wk.mdb";
// console.log('dbFilePath ===> : ', dbFilePath);

// fs.readFile(dbFilePath, (err, data) => {
//   if (err) {
//     console.error('เกิดข้อผิดพลาดในการอ่านไฟล์:', err);
//     return;
//   }

//   try {
//     const reader = new MDBReader(data);
//     const tableNames = reader.getTableNames();
//     console.log('รายชื่อตาราง:', tableNames);

//     const tableNameToCheck = 'member';
//     if (tableNames.includes(tableNameToCheck)) {
//       const table = reader.getTable(tableNameToCheck);
//       const columns = table.getColumnNames();
//       const rows = table.getData();

//       console.log(`ข้อมูลในตาราง ===> : ${tableNameToCheck}`);
//       console.log('header ===> :', columns);
//       rows.forEach(row => {
//         console.log(row);
//       });
//     } else {
//       console.log(`ไม่พบตารางชื่อ ${tableNameToCheck}`);
//     }

//   } catch (error) {
//     console.error('เกิดข้อผิดพลาดในการประมวลผลฐานข้อมูล:', error);
//   }
// });