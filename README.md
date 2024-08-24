# Excel Export for Yii2

Bu dokumentatsiya `yii2` ilovasida Excel fayllarni eksport qilish uchun ishlatiladigan funksiyani tushuntiradi.

## Tavsif

`ExcelExport` klassi `PhpOffice\PhpSpreadsheet` kutubxonasidan foydalanib, Excel fayllarini yaratadi va eksport qiladi. `actionExcelExport` metodida `BankMoneyLists` modelidan ma'lumotlar olinadi va Excel formatida saqlanadi.

## Kod

### `actionExcelExport` Metodi

```php
/**
 * Excel formatida eksport qilish.
 *
 * @throws Exception
 * @return string Excel fayl uchun yuklash havolasi
 */
public function actionExcelExport(): string
{
    $students = BankMoneyLists::find()->asArray()->all();

    $col = ['A', 'B', 'C', 'D', 'E'];

    $headers = ['id', 'user_id', 'branch_id', 'balance', 'counterparty'];

    return (new ExcelExport())->export($students, $headers, 'excel/', $col);
}
```
### ExcelExport class

```php
<?php

namespace app\models\helpers;

use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Yii;

class ExcelExport
{

/**
 * Excel faylini eksport qiladi.
 *
 * @param array $model Ma'lumotlar ro'yxati
 * @param array $headers Excel sarlavhalari
 * @param string $filePath Fayl saqlanadigan yo'l
 * @param array $col Excel ustunlari
 * @throws Exception
 * @return string Excel fayl uchun yuklash havolasi
 */
public function export(array $model, array $headers, string $filePath, $col): string
{
    $this->validateHeadersAndColumns($headers, $col);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $this->setHeaders($sheet, $headers, $col);
    $this->populateData($sheet, $model, $headers, $col);

    $this->createDirectory($filePath);
    $filePath = $this->getPhotoAlias($filePath);

    $this->saveFile($spreadsheet, $filePath);

    return $this->generateDownloadLink($filePath);
}

/**
 * Sarlavhalar va ustunlar sonini tekshiradi.
 *
 * @param array $headers Excel sarlavhalari
 * @param array $col Excel ustunlari
 * @throws InvalidArgumentException
 */
private function validateHeadersAndColumns(array $headers, array $col): void
{
    if (count($headers) !== count($col)) {
        throw new InvalidArgumentException('Headers and columns must have the same number of elements.');
    }
}


/**
 * Excel jadvalining sarlavhalarini o'rnatadi.
 *
 * @param $sheet PhpSpreadsheet sahifasi
 * @param array $headers Excel sarlavhalari
 * @param array $col Excel ustunlari
 */
private function setHeaders($sheet, array $headers, array $col): void
{
    foreach ($headers as $index => $header) {
        $cell = $col[$index] . '1';
        $sheet->setCellValue($cell, $header);
    }
}


/**
 * Excel jadvalini ma'lumotlar bilan to'ldiradi.
 *
 * @param $sheet PhpSpreadsheet sahifasi
 * @param array $model Ma'lumotlar ro'yxati
 * @param array $headers Excel sarlavhalari
 * @param array $col Excel ustunlari
 */
private function populateData($sheet, array $model, array $headers, array $col): void
{
    $row = 2;

    foreach ($model as $item) {
        foreach ($headers as $index => $header) {
            $column = $col[$index];
            $sheet->setCellValue($column . $row, $item[$header]);
        }

        $row++;
    }
}

/**
 * Fayl yo'lini alias bilan o'zgartiradi.
 *
 * @param string $filePath Fayl yo'li
 * @return string Alias bilan fayl yo'li
 */
private function getPhotoAlias($filePath): string
{
    return Yii::getAlias('@appWeb/' . $filePath . $this->generatePhotoName());
}


/**
 * Yangi Excel fayl nomini yaratadi.
 *
 * @return string Excel fayl nomi
 */
private function generatePhotoName(): string
{
    return (int)(microtime(true) * (1000)) . '.' . 'xlsx';
}


/**
 * Fayl saqlanadigan katalogni yaratadi.
 *
 * @param string $filePath Fayl yo'li
 */
private function createDirectory(string $filePath): void
{
    $directory = dirname($this->getPhotoAlias($filePath));
    if (!file_exists($directory)) {
        mkdir($directory, 0777, true);
    }
}

/**
 * Excel fayl uchun yuklash havolasini yaratadi.
 *
 * @param string $filePath Fayl yo'li
 * @return string Yuklash havolasi
 */
private function generateDownloadLink(string $filePath): string
{
    return "https://example/excel-export/" . basename($filePath);
}

}
```
