const { Builder, By, Key, until } = require('selenium-webdriver');
var Excel = require("exceljs");
const prompt = require('prompt-sync')();
var macaddress = require('macaddress');



// const username = prompt('Username: ');
// const password = prompt('Passowrd: ');

async function start() {
    var workbookr = new Excel.Workbook();
    await workbookr.xlsx.readFile("Tri.xlsx");
    var worksheetr = workbookr.getWorksheet(1);

    var lRow = worksheetr.lastRow;
    row = worksheetr.getRow(1);

    const workbookw = new Excel.Workbook();
    const worksheetw = workbookw.addWorksheet("My Sheet");

    let arr = [];
    for (let i = 1; i <= worksheetr.actualColumnCount; i++) {
        arr.push({ header: row.getCell(i).value, key: row.getCell(i).value, width: 15 })
        // console.log(worksheet.columns[i-1].header)
    }

    worksheetw.columns = arr

    let driver = await new Builder().forBrowser('chrome').build();

    try {
        await login(driver)
        var final = []
        let status, dob, result;
        for (i = 1; i < 2/*lRow._number*/; i++) {
            var data = {}
            row = worksheetr.getRow(i + 1);
            await inputPatientData(driver, row.getCell(5).value, row.getCell(4).value, row.getCell(3).value)
            // result = await checkEligibility(driver)
            // status = result[0];
            // dob = result[1];
            // data[worksheetw.columns[0].header] = status
            // data[worksheetw.columns[1].header] = dob
            // for (let j = 2, k = 3; j < (worksheetr.actualColumnCount); j++, k++) {
            //     data[worksheetw.columns[j].header] = row.getCell(k).value
            // }
            // final.push(data)
        }
        // console.log(final)
        // worksheetw.addRows(final)

        // save under export.xlsx
        // await workbookw.xlsx.writeFile('export.xlsx');
    } catch (err) {
        for (let j = 2, k = 3; j < (worksheetr.actualColumnCount); j++, k++) {
            data[worksheetw.columns[j].header] = row.getCell(k).value
        }
        worksheetw.addRows(final)

        // save under export.xlsx
        await workbookw.xlsx.writeFile('export-incomplete.xlsx');
        console.log("start: ", err);
    }
    finally {
        // await driver.quit();
    }
};


async function checkEligibility(driver) {
    return new Promise(async (resolve, reject) => {
        try {
            let status;
            let dob = 0;
            let html_text = await driver.wait(until.elementLocated(By.xpath("//*[@id='trnEligibilityStatus']")), 20000)
                .getAttribute('class')

            if (html_text == "stsRejected") {
                error = await driver.wait(until.elementLocated(By.xpath("//*[@id='eligErrorMsg']/span")), 20000).getText()
                error_text = error.split(" ");
                status = error_text[3] + " ERROR"
            }

            if (html_text == "stsinactive") {
                dob_text = await driver.wait(until.elementLocated(By.xpath("//*[@id='nadDOB']")), 20000).getText()
                dob = dob_text.split(" ")[1]
                try {
                    death_date = await driver.wait(until.elementLocated(By.xpath("//*[@id='BasicProfile']/dl/dt[3]")), 20000)
                    status = "DEAD"
                } catch (error) {
                    status = "inactive"
                }
            }
            

            if (html_text == "stsactive") {
                dob_text = await driver.wait(until.elementLocated(By.xpath("//*[@id='nadDOB']")), 20000).getText()
                dob = dob_text.split(" ")[1]

                try {
                    yellow_text = await driver.wait(until.elementLocated(By.xpath("//*[@id='other-payer-alert']")), 1000)
                    try {
                        benefit_info = await driver.wait(until.elementLocated(By.xpath("//*[@id='panel_benefitinformation']/a[6]")), 1000)

                        ppo = await driver.wait(until.elementLocated(By.xpath("//*[@id='BenefitsTable5']/table/tbody/tr[2]/td[3]")), 20000)
                            .getAttribute("innerHTML")
                        ppo_text = ppo.split(" ")

                        ppo_company_text = await driver.wait(until.elementLocated(By.xpath("//*[@id='BenefitsTable5']/table/tbody/tr[3]/td/dl[1]/dd")), 20000)
                            .getAttribute("innerHTML")

                        status = ppo_text[3] + ppo_company_text

                    } catch (error) {
                        status = "MSP"
                    }
                } catch (error) {
                    status = "MED B"
                }
            }

            resolve([status, dob])

        } catch (error) {
            console.log("checkEligibility: ", error)
            reject("Error in check elig");
        }

    })


};

async function inputPatientData(driver, ID, lastName, firstName) {
    return new Promise(async (resolve, reject) => {
        try {
            let eligibility_url = "https://mytools.gatewayedi.com/ManagePatients/RealTimeEligibility/Index"//"https://mytools.gatewayedi.com/ManagePatients/RealTimeEligibility/Request.aspx?payerid=00523"
            await driver.get(eligibility_url)

            await driver.wait(until.elementLocated(By.xpath("//*[@id='Medicare']/a")), 20000).click()
            await driver.wait(until.elementLocated(By.xpath("//*[@id='Medicare']/ul/li/a")), 20000).click()

            await driver.wait(until.elementLocated(By.xpath("//*[@id='EligibilityRequestPayerInquiry_EligibilityRequestFieldValues_SelectedProviderValue']/option[2]")), 20000).click()
            // await driver.wait(until.elementLocated(By.xpath('//*[@id="ProviderId"]/option[2]')), 20000).click()

            await driver.wait(until.elementLocated(By.xpath("//*[@id='EligibilityRequestPayerInquiry_EligibilityRequestFieldValues_SelectedSearchByValue']/option[3]")), 20000).click()
            // await driver.wait(until.elementLocated(By.xpath("//*[@id='QueryOptions']/option[3]")), 20000).click()
            
            
            if (ID && lastName && firstName) {
                await driver.findElement(By.name("EligibilityRequestTemplateInquiry.EligibilityRequestFieldValues.InsuranceNum")).sendKeys("username")
                reject("error")
                // await driver.findElement(By.name("EligibilityRequestTemplateInquiry.EligibilityRequestFieldValues.InsuranceNum")).sendKeys("username")
                // await driver.findElement(By.name("EligibilityRequestTemplateInquiry.EligibilityRequestFieldValues.InsuranceNum")).sendKeys("username")
                // await driver.findElement(By.name("EligibilityRequestTemplateInquiry.EligibilityRequestFieldValues.InsuranceNum")).sendKeys("username")
                // await driver.findElement(By.name("EligibilityRequestTemplateInquiry.EligibilityRequestFieldValues.InsuranceNum")).sendKeys("username")
                // await driver.findElement(By.name("EligibilityRequestTemplateInquiry.EligibilityRequestFieldValues.InsuranceNum")).sendKeys("username")
                
                // await driver.wait(until.elementLocated(By.xpath("//*[@id='EligibilityRequestTemplateInquiry_EligibilityRequestFieldValues_InsuranceNum']")), 20000).sendKeys("ID")
                // await driver.wait(until.elementLocated(By.xpath("//*[@id='EligibilityRequestTemplateInquiry_EligibilityRequestFieldValues_InsuredFirstName']")), 20000).sendKeys("firstName")
                // // await driver.wait(until.elementLocated(By.xpath("//*[@id='EligibilityRequestTemplateInquiry_EligibilityRequestFieldValues_InsuredLastName']")), 20000).sendKeys(lastName)
                // await driver.wait(until.elementLocated(By.xpath("//*[@id='btnUpload']")), 20000).click()
                // reject(new Error("Error in input data"))
                //*[@id='btnUploadButton']/span/span
                console.log("IDD: ", ID, lastName, firstName)
            }

            // if (ID && lastName && firstName) {
            //     await driver.wait(until.elementLocated(By.xpath("//*[@id='MemberID']")), 20000).sendKeys(ID)
            //     await driver.wait(until.elementLocated(By.xpath("//*[@id='Subscriber_Last_Name']")), 20000).sendKeys(lastName)
            //     await driver.wait(until.elementLocated(By.xpath("//*[@id='Subscriber_First_Name']")), 20000).sendKeys(firstName, Key.RETURN)
            // }
            resolve(1)
        } catch (error) {
            console.log("inputPatientData: ", error)
            reject(new Error("Error in input data"))
        }
    })
};


async function login(driver) {
    return new Promise(async (resolve, reject) => {
        try {
            
            let username = '2wt52';
            let password = 'Wyoming@17702022';
            await driver.get("https://mytools.gatewayedi.com/LogOn");
            await driver.findElement(By.name("UserName")).sendKeys(username)//.sendKeys(username);
            await driver.findElement(By.name("Password")).sendKeys(password, Key.RETURN)//.sendKeys(password, Key.RETURN);
            resolve(1)
            
        } catch (error) {
            reject(new Error("Error in Login"))
        }
    })
};


macaddress.one('Ethernet').then(function (mac) {
    if (mac == "70:5a:0f:cf:93:f8") {
        start();
    }
    // if (mac == "18:03:73:c6:84:1c") {
    //     start();
    // }
    else {
        console.log("not applicable");
    }
});