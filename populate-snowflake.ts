
/** 
 * This is supposed to work in an Excel O365 sheet
 * Example file should be like this:
 *  |A       |B          |
 * 0|SITE_ID |SITE_NAME  |
 * 1|1       |Bogota     |
 * 2|2       |Medellin   |
 * 3|3       |Santa Marta|
 * 
 * It will insert that data in the configured Snowflake instance
 */
async function main(workbook: ExcelScript.Workbook) {
    try {
      const sites = getTableData(workbook.getActiveWorksheet());
      console.log('REMOVING DATA FROM TABLE ...');
      await removeDataFromTable();
      console.log('INSERTING SITES ...');
      const responses = await insertSites(sites);
      console.log(`INSERTED ${responses.length} SITES`);
    } catch (error){
      throw new Error(error.message);
    }  
  }
  
  function getTableData(sheet: ExcelScript.Worksheet) {
    const data: Record<number, string> = {};
    const usedRange = sheet.getUsedRange();
    const siteIdData = getColumnData(usedRange, 0);
    const siteNameData = getColumnData(usedRange, 1);
    
    if (siteIdData.length !== siteNameData.length) {
      throw new Error(`SITE_ID and SITE_NAME columns don't have matching rows. SITE_ID row count: ${siteIdData.length}  SITE_NAME row count: ${siteNameData.length}`);
    }
  
    for (let i = 1; i < siteIdData.length; i++) {
      data[parseInt(siteIdData[i])] = siteNameData[i];
    }
  
    return data;
  }
  
  function getColumnData(usedRange: ExcelScript.Range, idx: number) {
    const data: string[] = [];
  
    const column = usedRange.getColumn(idx);
    const rowCount = column.getRowCount();
    const dataValues = column.getValues();
  
    for (let row = 0; row < rowCount; row++) {
      const value = dataValues[row].toString();
      if(value !== "") {
        data.push(dataValues[row].toString());
      }    
    }
  
    return data;
  }
  
  function getRequestHeaders(): Record<string, string> {
    return {
      'Content-Type': 'application/json',
      'User-Agent': 'PopulateSnowflake/0.1',
      'Authorization': 'Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJPWFZMWFpFLVFJQjMzNDM4LlJCQVJPTkMuU0hBMjU2OnJwL09jb1Q3dWVTejZXUjVBSm15c0dzM2tBa0lHSmtNZW93L1M2UllmNWs9Iiwic3ViIjoiT1hWTFhaRS1RSUIzMzQzOC5SQkFST05DIiwiZXhwIjoxNzEyMjkxMzc5LCJpYXQiOjE3MTIyODc3Nzl9.otYbO4CZpCLHnVGtemY_GIpWMbFzVbHvMMU9YVN_iLmanFsL2iqJGvU3Id3-ch-ladcNrk8NxV4CvbmEWhFwptLaI-Ht5z6NcSGA5a8X64Mx8BGGDQIbRQP7Sv-JD9tYbf8p6Q0bhZf0zO2_bXF6b0ZCwzOZhNPMU1GdDVVyiLqJZQQI6zjLI1y9mkSouV0VcyySFKGyrhEG8sziBJbLhEuX369IF2QLV0xhzUXlJBtLMfu-Ub1vehkGPKQJZTw_eHUzzqGtlKN09w16mvqWG9zcDttfUaIbWtcvLt_nw1cyV4Vs_Togz-QMcC75oqGofREZfTuHAga0DfoKKCxAew',
      'Accept': 'application/json',
      'X-Snowflake-Authorization-Token-Type': 'KEYPAIR_JWT' 
    }
  }
  
  function runSnowflakeStatement(statement: string): Promise < Response >{
    return fetch('https://yvb65666.us-east-1.snowflakecomputing.com/api/v2/statements', {
      method: "POST",
      headers: getRequestHeaders(),
      body: JSON.stringify({
        "statement": statement,
        "timeout": 60,
        "database": "O365_EXAMPLE",
        "schema": "SITES",
        "warehouse": "TEST_WH"
      })
    });
  }
  
  function removeDataFromTable()  {
    return runSnowflakeStatement("DELETE FROM SITES");
  }
  
  function insertSites(sites: Record<number, string>): Promise<Response[]> {
    const siteNames = Object.values(sites);
    const siteIds = Object.keys(sites);
    const sitesLength = siteIds.length;
    const promises: Promise<Response>[] = [];
    
    for(let i = 0; i < sitesLength; i++) {
      promises.push(runSnowflakeStatement(`INSERT INTO SITES VALUES(${siteIds[i]}, '${siteNames[i]}')`));
    }
  
    return Promise.all(promises);
  }
  