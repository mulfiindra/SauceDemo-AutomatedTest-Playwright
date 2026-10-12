
import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import path from 'node:path';

async function readLoginData(): Promise<{
  username: string;
  password: string;
}> {
  const filePath = path.resolve(
    process.cwd(),
    'test-data',
    'SauceDemo_Credential.xlsx'
  );

  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.getWorksheet('LoginData');

  if (!worksheet) {
    throw new Error(
      'Worksheet "LoginData" was not found.'
    );
  }

  // Row 1 = header, Row 2 = login data
  const row = worksheet.getRow(2);

  const username = row.getCell(1).text.trim();
  const password = row.getCell(2).text.trim();

  if (!username || !password) {
    throw new Error(
      'Username or password is empty in the Excel file.'
    );
  }

  return { username, password };
}

test('TC-01 - Failed login with invalid credentials', async ({
  page,
}) => {
  // 1. Open SauceDemo login page
  await page.goto('https://www.saucedemo.com/');

  // 2. Enter invalid username
  await page.locator('[data-test="username"]').fill('invaliduser');

  // 3. Enter invalid password
  await page.locator('[data-test="password"]').fill('invalidpassword');

  // 4. Click Login button
  await page.locator('[data-test="login-button"]').click();

  // 5. Verify the error message
  const errorMessage = page.locator('[data-test="error"]');

  await expect(errorMessage).toBeVisible();

  await expect(errorMessage).toHaveText(
    'Epic sadface: Username and password do not match any user in this service'
  );
});

test('TC-02 - Successful login with hardcoded credentials', async ({
  page,
}) => {
  console.log('========================================');
  console.log('TC-02 - SUCCESSFUL LOGIN HARDCODED');
  console.log('========================================');

  // 1. Open SauceDemo login page
  await page.goto('https://www.saucedemo.com/');

  // 2. Enter valid username
  await page.locator('[data-test="username"]').fill('standard_user');

  // 3. Enter valid password
  await page.locator('[data-test="password"]').fill('secret_sauce');

  // 4. Click Login button
  await page.locator('[data-test="login-button"]').click();

  console.log('Username  : standard_user');
  console.log('Login     : CLICKED');

  // 5. Verify successful navigation
  await expect(page).toHaveURL(/\/inventory\.html$/);

  // 6. Verify inventory page is visible
  await expect(page.locator('.inventory_list')).toBeVisible();

  console.log('Current URL:', page.url());
  console.log('Assertion : PASS');
});


test('TC-03 - Successful login with Excel data', async ({
  page,
}) => {
  console.log('========================================');
  console.log('TC-03 - SUCCESSFUL LOGIN FROM EXCEL');
  console.log('========================================');

  // 1. Read credentials from Excel
  const loginData = await readLoginData();

  console.log('Excel Username:', loginData.username);

  // 2. Open SauceDemo and perform login
  await page.goto('https://www.saucedemo.com/');

  await page.locator('[data-test="username"]')
    .fill(loginData.username);

  await page.locator('[data-test="password"]')
    .fill(loginData.password);

  await page.locator('[data-test="login-button"]').click();

  console.log('Login     : CLICKED');

  // 3. Verify successful navigation
  await expect(page).toHaveURL(/\/inventory\.html$/);

  // 4. Verify inventory content
  await expect(page.locator('.inventory_list')).toBeVisible();

  console.log('Current URL:', page.url());
  console.log('Assertion : PASS');
});