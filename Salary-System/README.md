# FOIS Salary System

Новая архитектура расчёта зарплат водителей.

## Главный принцип

- Один водительский лист = текущий расчёт одного водителя.
- Одна строка в `Salary Database` = один водитель за одну неделю.
- `Week Key`: `2026-W30`.
- `Record Key`: `2026-W30|709`.
- Исторические записи после архивирования не должны зависеть от живых формул водительского листа.

## Листы

| Лист | Назначение |
|---|---|
| `README` | Инструкция по работе с книгой |
| `Settings` | Driver ID, водитель, truck, тип оплаты, ставка, dispatcher, recruiter, status |
| `Driver Template` | Стандартный шаблон расчёта одного водителя |
| `D_709`, `D_715` | Примеры водительских листов |
| `Salary Database` | Накопительная база зарплат |
| `Salary Report` | Отчёт по выбранному `Week Key` |
| `Archive Log` | Журнал архивирования завершённых расчётов |

## Структура записи Salary Database

```text
Year
Week
Week Key
Driver ID
Driver Name
Truck
Pay Type
Rate
Loaded Miles
Empty Miles
Total Miles
Gross
Adjusted Gross
Driver Pay
Deductions
Net Pay
Record Key
Status
```

## Логика оплаты

```gs
=IF(PayType="CPM",TotalMiles*Rate,IF(PayType="% Gross",AdjustedGross*Rate,0))
```

## Отчёт по неделе

В `Salary Report!B4` выбирается `Week Key`, например:

```text
2026-W30
```

Отчёт фильтрует `Salary Database` по этому ключу.

## Рабочий процесс

1. Добавить водителя в `Settings`.
2. Скопировать `Driver Template`.
3. Назвать новый лист `D_<Driver ID>`.
4. Заполнить год, неделю и рейсы.
5. Проверить итоговый блок.
6. После закрытия недели перенести итоговую строку в `Salary Database` как значения.
7. Проверить `Salary Report`.
8. После тестового периода автоматизировать архивирование через Apps Script.

## Файлы

- `Code.gs` — скрипт для создания структуры в новой Google Sheets книге.
- Готовый Excel-шаблон создан отдельно как `FOIS_Salary_System_2026.xlsx`.

## Следующий этап

После нескольких проверенных недель:

- добавить кнопку `Archive Week`;
- запретить дублирование `Record Key`;
- перенести старую историю;
- подключить Weekly Report;
- подготовить данные к PostgreSQL и Power BI.
