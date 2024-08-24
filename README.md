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

