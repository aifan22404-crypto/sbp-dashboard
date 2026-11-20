# สุ่มตรวจตัวเลขในบทความ (20 พ.ย. 2569)

เทียบตัวเลขที่อ้างในบทความกับ Master Panel และตาราง Power BI ด้วยสคริปต์ `scripts/spot_check_article_2026-11-20.py` (ยอมคลาดเคลื่อนจากการปัดเศษ 0.5%)

| ข้อความในบทความ                              |      ค่าในบทความ |     ค่าจากข้อมูล | ผล   |
|:---------------------------------------------|-----------------:|-----------------:|:-----|
| Poverty Pattani 2568 (%)                     |     21.21        |     21.21        | ตรง  |
| Poverty Whole Kingdom 2568 (%)               |      3.27        |      3.27        | ตรง  |
| Poverty Pattani 2564 (%)                     |     30.43        |     30.43        | ตรง  |
| Poverty Yala 2566 (%)                        |      9.04        |      9.04        | ตรง  |
| Poverty Yala 2567 (%)                        |     25.54        |     25.54        | ตรง  |
| LQ agriculture Pattani                       |      2.69        |      2.69        | ตรง  |
| LQ agriculture Yala                          |      4.14        |      4.14        | ตรง  |
| LQ manufacturing Yala                        |      0.29        |      0.29        | ตรง  |
| LQ manufacturing Pattani                     |      0.36        |      0.36        | ตรง  |
| LQ education Narathiwat                      |      4           |      4           | ตรง  |
| Fish landings Pattani 2567 (t)               |  84694           |  84694           | ตรง  |
| Fish landings value Pattani 2567 (mb)        |   3359           |   3359.3         | ตรง  |
| Manufacturing share of GPP Pattani 2567 (%)  |      8.72        |      8.72        | ตรง  |
| Visitors Narathiwat 2565                     | 431003           | 431003           | ตรง  |
| Visitors Yala 2568                           |      2.03327e+06 |      2.03327e+06 | ตรง  |
| Malaysian hotel guests Pattani 2568          |  54926           |  54926           | ตรง  |
| Occupancy Yala 2568 (%)                      |     80.9         |     80.93        | ตรง  |
| GPP per capita Narathiwat 2567 (baht)        |  66785           |  66785           | ตรง  |
| GPP per capita Songkhla 2567 (baht)          | 156755           | 156755           | ตรง  |
| HH income national 2566 (baht/month)         |  29030           |  29030           | ตรง  |
| Informal debt Pattani 2566 (baht)            |  11847           |  11847           | ตรง  |
| Unemployment Narathiwat 2564 (%)             |      8.47        |      8.47        | ตรง  |
| Unemployment Narathiwat 2568 (%)             |      2.37        |      2.37        | ตรง  |
| Official incidents Narathiwat 2568           |    103           |    103           | ตรง  |
| UCDP lethal events Narathiwat 2568           |     13           |     13           | ตรง  |
| Vessels Pattani 2564                         |   2214           |   2214           | ตรง  |
| Vessels Pattani 2568                         |   1624           |   1624           | ตรง  |
| Border trade Narathiwat 2564 (mb)            |   3405           |   3404.89        | ตรง  |
| Manufacturing dissolved per 100 new (3 prov) |     50.4         |     50.4         | ตรง  |
| RSS3 annual average 2554 (baht/kg)           |    127.96        |    127.96        | ตรง  |
| RSS3 annual average 2561 (baht/kg)           |     41.48        |     41.48        | ตรง  |
| RSS3 months at or above 100 baht             |     14           |     14           | ตรง  |
| Chanae lethal events 2564-2568               |     13           |     13           | ตรง  |
| Nong Chik new firms 2565-2568                |     64           |     64           | ตรง  |

สรุป: {'ตรง': 34}

หมายเหตุ: นี่คือการตรวจว่าบทความคัดลอกตัวเลขจากชุดข้อมูลถูกต้อง ไม่ใช่การตรวจชุดข้อมูลกับไฟล์ต้นฉบับของหน่วยงาน (ส่วนนั้นตรวจตอนทำ clean และบันทึกใน red flags)
