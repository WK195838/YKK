# WorkFlow 與 VS2005 企業系統 - 系統規格書

## 1. 系統概述

### 1.1 系統簡介
本規格書涵蓋 YKK 企業的兩大核心業務系統：
- **WorkFlow 系統**：VS2003 版本的企業工作流程管理系統
- **VS2005 系統群**：包含 24 個子系統的大型企業應用程式群組

### 1.2 技術概要
| 系統 | 技術棧 | 檔案數量 | 主要功能 |
|------|--------|----------|----------|
| WorkFlow | ASP.NET 1.1 + VB.NET + SQL Server | 920個 | 製造業流程管理 |
| VS2005 | ASP.NET 2.0 + VB.NET + SQL Server | 2,762個 | 企業管理系統群 |

## 2. WorkFlow 系統詳細分析

### 2.1 系統架構

```mermaid
graph TB
    subgraph "WorkFlow 系統架構"
        A[用戶介面層<br/>199個ASPX頁面] --> B[業務邏輯層<br/>207個VB.NET類別]
        B --> C[資料存取層<br/>SQL Server連接]
        C --> D[資料庫層<br/>WorkFlow + SFD]
    end
    
    subgraph "檔案系統"
        E[文件管理<br/>43個目錄] --> F[工作表範本<br/>113個Excel檔案]
        F --> G[流程圖表<br/>25個工作表目錄]
    end
    
    A --> E
    
    subgraph "核心模組"
        H[製造進料<br/>ManufIn]
        I[製造出料<br/>ManufOut]
        J[表面處理<br/>Surface]
        K[進口作業<br/>Import]
        L[流程地圖<br/>Map]
        M[系統管理<br/>System]
    end
    
    B --> H
    B --> I
    B --> J
    B --> K
    B --> L
    B --> M
```

### 2.2 資料庫架構

```mermaid
erDiagram
    WorkFlow ||--o{ T_Form : "表單管理"
    WorkFlow ||--o{ T_FormData : "表單資料"
    WorkFlow ||--o{ T_WaitHandle : "待處理事項"
    WorkFlow ||--o{ T_User : "用戶管理"
    WorkFlow ||--o{ T_Process : "流程定義"
    
    SFD ||--o{ Manufacturing : "製造資料"
    SFD ||--o{ Import : "進口資料"
    SFD ||--o{ Surface : "表面處理"
    
    T_Form {
        string FormNo "表單號碼"
        int FormSno "流水號"
        string ApplyID "申請者"
        datetime CreateTime "建立時間"
        int Status "狀態"
    }
    
    T_FormData {
        string FormNo "表單號碼"
        int FormSno "流水號"
        string FieldName "欄位名稱"
        string FieldValue "欄位值"
    }
```

### 2.3 檔案結構分析

#### 程式檔案 (610個，66.3%)
| 檔案類型 | 數量 | 用途 |
|----------|------|------|
| .aspx | 199 | 用戶介面頁面 |
| .vb | 207 | 業務邏輯程式碼 |
| .resx | 200 | 資源檔案 |
| .config | 4 | 系統配置 |

#### 非程式檔案 (310個，33.7%)
| 檔案類型 | 數量 | 用途 |
|----------|------|------|
| .xls/.xlsx | 113 | Excel範本與流程圖 |
| 圖像檔案 | 165 | UI圖示與圖表 |
| 文件檔案 | 8 | 說明文件 |
| 其他 | 24 | 專案檔案與其他 |

### 2.4 核心功能模組

```mermaid
graph LR
    subgraph "製造模組"
        A[ManufIn<br/>製造進料]
        B[ManufOut<br/>製造出料]
        C[Surface<br/>表面處理]
    end
    
    subgraph "貿易模組"
        D[Import<br/>進口作業]
        E[AppendSpec<br/>規格追加]
        F[ColorAppend<br/>顏色追加]
    end
    
    subgraph "系統模組"
        G[NewSystem<br/>新系統申請]
        H[ChangeSystem<br/>系統變更]
        I[ChangeData<br/>資料變更]
        J[TroubleShooter<br/>問題處理]
    end
    
    subgraph "流程管理"
        K[Map<br/>流程地圖]
        L[MapMod<br/>流程修改]
        M[Sample<br/>樣品管理]
    end
```

### 2.5 工作流程

```mermaid
sequenceDiagram
    participant U as 申請者
    participant S as 系統
    participant A as 審核者
    participant D as 資料庫
    
    U->>S: 1. 提交表單申請
    S->>D: 2. 儲存表單資料
    S->>A: 3. 發送審核通知
    A->>S: 4. 執行審核作業
    S->>D: 5. 更新審核狀態
    S->>U: 6. 回傳審核結果
    
    alt 審核通過
        S->>S: 7. 進入下一流程
    else 審核駁回
        S->>U: 8. 要求修正重送
    end
```

## 3. VS2005 系統群詳細分析

### 3.1 系統架構總覽

```mermaid
graph TB
    subgraph "VS2005 企業系統群架構"
        A[Web前端層<br/>633個ASPX頁面] --> B[業務邏輯層<br/>713個VB.NET類別]
        B --> C[資料存取層<br/>多資料庫連接]
        C --> D[資料庫群<br/>24個業務資料庫]
    end
    
    subgraph "系統分類"
        E[財務系統<br/>N2W]
        F[貿易系統<br/>EDI/DTMW]
        G[品質系統<br/>ISOS/ISOSQC]
        H[物料系統<br/>ItemRegister]
        I[工作流<br/>WorkFlowSub]
    end
    
    B --> E
    B --> F
    B --> G
    B --> H
    B --> I
```

### 3.2 主要子系統分析

#### 3.2.1 N2W (報銷系統) - 120 個頁面
```mermaid
graph LR
    subgraph "N2W 報銷系統"
        A[出差申請<br/>BusinessTrip] --> B[費用申請<br/>Expense]
        B --> C[折扣申請<br/>Discount]
        C --> D[資金申請<br/>Funding]
        D --> E[結帳處理<br/>CloseAccount]
        E --> F[客訴處理<br/>Complaint]
    end
```

**核心功能：**
- 出差費用管理 (BusinessTripSheet_01-03.aspx)
- 費用核銷處理 (ExpenseSheet_01-03.aspx)
- 折扣管理 (DiscountSheet_01-02.aspx)
- 資金申請 (FundingSheet_01-03.aspx)
- 結帳作業 (CloseAccountSheet_01-03.aspx)

#### 3.2.2 EDI (電子資料交換) - 68 個頁面
```mermaid
graph LR
    subgraph "EDI 系統"
        A[訂單處理<br/>Order] --> B[發貨通知<br/>Shipment]
        B --> C[發票處理<br/>Invoice]
        C --> D[庫存同步<br/>Inventory]
        D --> E[物料展開<br/>Material]
    end
```

**系統特色：**
- 與 WAVES 系統整合
- 支援多種 EDI 格式 (TNF, ADIDAS, REEBOK)
- 自動物料展開功能
- 批次處理與監控

#### 3.2.3 DTMW (染料管理) - 65 個頁面
專門用於染料相關的製造流程管理。

#### 3.2.4 ISOSQC (ISO品質控制) - 61 個頁面
實施 ISO 品質管理標準的專業系統。

### 3.3 檔案統計分析

#### 總體統計
- **總檔案數：** 2,762 個
- **程式檔案：** 1,377 個 (49.9%)
- **非程式檔案：** 1,385 個 (50.1%)

#### 檔案類型分佈
| 類型 | 數量 | 百分比 | 用途 |
|------|------|--------|------|
| .vb | 713 | 25.8% | VB.NET 程式碼 |
| .aspx | 633 | 22.9% | Web 頁面 |
| 圖像檔案 | 686 | 24.8% | UI 圖示與圖表 |
| .xls/.xlsx | 233 | 8.4% | Excel 範本 |
| .dll | 123 | 4.5% | 程式庫 |
| .config | 31 | 1.1% | 設定檔案 |
| 其他 | 343 | 12.5% | 文件與其他 |

### 3.4 系統整合架構

```mermaid
graph TB
    subgraph "外部系統整合"
        A[WAVES 系統<br/>10.245.1.2] 
        B[EDI服務<br/>10.245.0.153]
        C[WFS服務<br/>10.245.1.50]
    end
    
    subgraph "VS2005 核心系統"
        D[N2W 報銷]
        E[EDI 貿易]
        F[DTMW 染料]
        G[ItemRegister 物料]
        H[ISOSQC 品質]
    end
    
    subgraph "資料庫群"
        I[EDI DB<br/>10.245.0.112]
        J[各系統專用DB]
        K[共用資料庫]
    end
    
    A --> E
    B --> E
    C --> E
    
    D --> I
    E --> I
    F --> J
    G --> J
    H --> J
```

## 4. 技術架構比較

### 4.1 技術棧對比
| 項目 | WorkFlow | VS2005 |
|------|----------|---------|
| .NET Framework | 1.1 | 2.0 |
| 開發工具 | VS2003 | VS2005 |
| 語言 | VB.NET | VB.NET |
| 資料庫 | SQL Server | SQL Server |
| 認證方式 | Windows | Windows |
| 部署方式 | IIS 6.0 | IIS 6.0 |

### 4.2 系統複雜度分析

```mermaid
pie title 系統複雜度分佈
    "N2W報銷" : 120
    "EDI貿易" : 68
    "DTMW染料" : 65
    "WorkFlowSub" : 63
    "ItemRegister物料" : 62
    "ISOSQC品質" : 61
    "其他系統" : 194
```

## 5. 部署架構

### 5.1 網路架構

```mermaid
graph TB
    subgraph "DMZ 區域"
        A[Web Server<br/>IIS 6.0]
        B[應用伺服器<br/>WorkFlow/VS2005]
    end
    
    subgraph "內部網路"
        C[資料庫伺服器<br/>10.245.1.20]
        D[EDI 資料庫<br/>10.245.0.112]
        E[WAVES 系統<br/>10.245.1.2]
        F[檔案伺服器<br/>10.245.0.192]
    end
    
    subgraph "服務層"
        G[EDI 服務<br/>10.245.0.153]
        H[WFS 服務<br/>10.245.1.50]
    end
    
    A --> B
    B --> C
    B --> D
    B --> E
    B --> F
    B --> G
    B --> H
```

### 5.2 資料流架構

```mermaid
sequenceDiagram
    participant U as 用戶
    participant W as Web層
    participant A as 應用層
    participant D as 資料庫
    participant E as 外部系統
    
    U->>W: HTTP請求
    W->>A: 業務處理
    A->>D: 資料查詢/更新
    A->>E: 外部系統調用
    E-->>A: 回傳結果
    D-->>A: 回傳資料
    A-->>W: 處理結果
    W-->>U: HTTP回應
```

## 6. 安全架構

### 6.1 認證與授權

```mermaid
graph LR
    subgraph "認證層"
        A[Windows認證]
        B[AD整合]
        C[角色管理]
    end
    
    subgraph "授權層"
        D[功能權限]
        E[資料權限]
        F[IP限制]
    end
    
    subgraph "安全層"
        G[SQL注入防護]
        H[XSS防護]
        I[檔案上傳限制]
    end
    
    A --> D
    B --> E
    C --> F
    D --> G
    E --> H
    F --> I
```

### 6.2 資料安全
- **資料庫加密：** 使用 SQL Server 內建加密
- **傳輸加密：** SSL/HTTPS 保護
- **存取控制：** Windows 認證 + 角色管理
- **稽核記錄：** 完整的操作日誌

## 7. 效能與擴展性

### 7.1 效能指標
| 系統 | 同時用戶數 | 回應時間 | 資料量 |
|------|------------|----------|--------|
| WorkFlow | 50-100 | < 3秒 | 中等 |
| VS2005 | 200-500 | < 2秒 | 大量 |

### 7.2 擴展策略

```mermaid
graph TB
    subgraph "水平擴展"
        A[負載平衡器] --> B[Web Server 1]
        A --> C[Web Server 2]
        A --> D[Web Server N]
    end
    
    subgraph "垂直擴展"
        E[CPU升級]
        F[記憶體擴充]
        G[儲存擴展]
    end
    
    subgraph "資料庫擴展"
        H[讀寫分離]
        I[分片策略]
        J[快取層]
    end
```

## 8. 維護與監控

### 8.1 系統監控

```mermaid
graph LR
    subgraph "監控指標"
        A[系統效能]
        B[資料庫效能]
        C[網路狀態]
        D[錯誤率]
    end
    
    subgraph "監控工具"
        E[IIS日誌]
        F[SQL Profiler]
        G[事件檢視器]
        H[自訂監控]
    end
    
    A --> E
    B --> F
    C --> G
    D --> H
```

### 8.2 備份策略
- **資料庫備份：** 每日完整備份 + 每小時增量備份
- **檔案備份：** 每週完整備份
- **設定備份：** 變更時即時備份
- **異地備援：** 重要資料異地存放

## 9. 升級建議

### 9.1 技術債務
1. **框架老舊：** .NET Framework 1.1/2.0 已停止支援
2. **安全風險：** 老舊技術存在安全漏洞
3. **維護困難：** 開發人員稀少
4. **擴展限制：** 無法利用現代化功能

### 9.2 現代化路徑

```mermaid
graph LR
    subgraph "階段一：穩定化"
        A[現況評估] --> B[風險分析]
        B --> C[優先級排序]
    end
    
    subgraph "階段二：升級"
        D[.NET Core 遷移] --> E[資料庫現代化]
        E --> F[UI/UX 重構]
    end
    
    subgraph "階段三：最佳化"
        G[雲端部署] --> H[微服務架構]
        H --> I[DevOps整合]
    end
    
    C --> D
    F --> G
```

### 9.3 建議技術棧
- **後端：** .NET 6/8 + C#
- **前端：** React/Vue.js + TypeScript
- **資料庫：** SQL Server 2022 + Redis
- **部署：** Docker + Kubernetes
- **雲端：** Azure/AWS

## 10. 總結

### 10.1 系統評估
| 項目 | WorkFlow | VS2005 | 評分 |
|------|----------|---------|------|
| 功能完整性 | ★★★★☆ | ★★★★★ | 4.5/5 |
| 技術先進性 | ★★☆☆☆ | ★★☆☆☆ | 2/5 |
| 維護難度 | ★★★☆☆ | ★★★★☆ | 3.5/5 |
| 安全性 | ★★★☆☆ | ★★★☆☆ | 3/5 |
| 擴展性 | ★★☆☆☆ | ★★★☆☆ | 2.5/5 |

### 10.2 關鍵發現
1. **業務價值高：** 兩系統都是企業核心業務支援系統
2. **技術債務重：** 使用過時技術，存在安全與維護風險
3. **整合複雜：** 系統間依賴關係複雜，升級需謹慎規劃
4. **資料豐富：** 累積大量業務資料，遷移需完整規劃

### 10.3 建議行動
1. **短期：** 安全性修補、效能優化、監控強化
2. **中期：** 逐步遷移至 .NET Core、資料庫升級
3. **長期：** 微服務重構、雲端化部署、現代化UI

---

**文件版本：** 1.0  
**建立日期：** 2025/01/24  
**更新日期：** 2025/01/24  
**文件狀態：** 初版完成 