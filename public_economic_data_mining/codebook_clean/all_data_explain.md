# data_clean 資料集用途總表

| 檔名 | 欄位重點 | 分析用途與可串接主題 |
| --- | --- | --- |
| `electricity_sales_by_usage.csv` | 年度、用途別、售電量 (kWh) | 長期用電結構：住宅 vs. 工業 vs. 商業等，占比變化、需求成長動能，可支援主題 2、6。 |
| `electricity_consumption_by_sector.csv` | Gregorian 月份、ROC 原始日期、31 個部門/產業用電 (MWh) | 高頻需求面：工業細項與服務業、運輸、農業等，可分析季節性、景氣循環、用電與 GDP/氣候的連動，對應主題 2、6、7。 |
| `utility_retail_statistics.csv` | 年度、電燈/電力售電量、用戶數、平均電價 | 公用售電業統計：客戶結構、用電合計與電價，連結成本或政策評估，對應主題 2、5。 |
| `generation_by_energy_type.csv` | 年度、發購電來源 (台電/購電)、能源別、發購電量 (kWh) | 供給側能源結構：燃煤/燃氣/再生能源等占比演進，可量化 CO₂ intensity 或能源轉型，支援主題 1、4、8、9、10。 |
| `generation_capacity_long.csv` | 月份、擁有者 (台電/IPP 等)、技術別、裝置容量 (MW) | 裝置容量時序：追蹤火力 vs. 再生能源的建置與退役，估算 capacity factor，支援主題 1、4、8、10。 |
| `generation_fuel_consumption_long.csv` | 月份、擁有者、燃料別、耗用量 (公噸油當量/10⁷千卡) | 燃料耗用：觀察燃煤、燃氣、再生能源燃料需求，評估能源安全與進口依賴，對應主題 1、8、9。 |
| `power_supply_demand_summary.csv` | 日期、尖峰供電能力、尖峰負載、備轉容量/率、工業與民生用電 (GWh) | 日度供需平衡：即時檢視尖峰負載是否逼近供電能力，配合氣候事件分析，支援主題 3、7、10。 |
| `power_unit_available_capacity.csv` | 日期、個別機組名稱、可用容量 (MW) | 機組可用量：追蹤重要火力/核能/再生機組在尖峰時段的表現，以找出備轉緊張原因，支援主題 3。 |
| `current_year_daily_reserve_margin.csv` | 日期、備轉容量 (MW)、備轉率 (%) | 當年度每日備轉指標：評估近期供電風險與季節性差異，支援主題 3、7。 |
| `three_year_daily_reserve_margin.csv` | 同上（三年歷史） | 提供近三年每日備轉比較，找出熱浪/寒流下的系統壓力，支援主題 3、7。 |
| `annual_peak_load_and_reserve.csv` | 年度尖峰負載 (MW)、備用容量率 (%) | 長期年度尖峰紀錄：用於宏觀供給能力 vs. 需求成長分析，支援主題 3、10。 |
| `solar_feed_in_records.csv` | ROC 年、月份 (含年總和旗標)、度數 (MWh) | 太陽光電購電實績：衡量再生能源成長速度與淡旺季產能，支援主題 4。 |
| `generation_costs_by_technology.csv` | 類別/子類別/項目、年度、單位成本 (元/度) | 2022–2024 發購電成本：比較火力 vs. 再生等成本結構，作為電價與政策分析基礎，支援主題 5、10。 |
| `nuclear_generation_performance.csv` | 年度、核能發電量 (GWh)、容量因數、減碳量 (噸) | 核能績效：評估核能在供電與減碳上的貢獻，支援主題 8、9。 |
| `nuclear_facility_radiation.csv` | 站名、站號、日期時間、劑量率、經緯度 | 核設施輻射監測：環境安全與社會信任指標，可輔助主題 8、9 或風險溝通。 |

> 備註：所有檔案皆由 `scripts/clean_data.py` 根據 `data/` 原始資料清洗後得到，欄位與單位已統一，可直接套用於 `target.txt` 所列 10 大分析主題。
