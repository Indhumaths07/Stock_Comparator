export interface UploadedFile {
  file: File;
  name: string;
  size: number;
  sheets: string[];
  selectedSheet: string;
  headers: string[];
  rawRows: Record<string, any>[];
  headerRowIndex?: number;
}

export type MatchMode = 'toon_label' | 'toon_and_style' | 'style_no' | 'ean';

export interface ColumnMapping {
  matchMode: MatchMode;
  skuColumn?: string; // EAN / Barcode fallback
  nameColumn: string; // Product Name
  wondersoftQtyColumn: string;
  ecommerceQtyColumn: string;
  mainCategoryColumn?: string;
  productTypeColumn?: string;
  wondersoftStoreNameColumn?: string;
  ecommerceStoreNameColumn?: string;
  wondersoftSkuColumn?: string;
  ecommerceSkuColumn?: string;
  wondersoftNameColumn?: string;
  ecommerceNameColumn?: string;
  wondersoftToonLabelColumn?: string;
  ecommerceToonLabelColumn?: string;
  toonLabelColumn?: string;
  wondersoftStyleNoColumn?: string;
  ecommerceStyleNoColumn?: string;
  styleNoColumn?: string;
  wondersoftColorColumn?: string;
  ecommerceColorColumn?: string;
  colorColumn?: string;
  wondersoftSizeColumn?: string;
  ecommerceSizeColumn?: string;
  sizeColumn?: string;
  priceColumn?: string;
  storeNameColumn?: string;
  columnExtraColumn?: string;
  storeComparisonMode?: 'all' | 'exact_store' | 'filter_store';
  selectedStore?: string;
}

export type StockStatus =
  | 'out_of_stock'
  | 'stock_decreased'
  | 'in_sync'
  | 'stock_increased'
  | 'only_in_wondersoft'
  | 'only_in_ecommerce';

export interface ComparisonItem {
  id: string;
  sku: string; // Display identifier (Style No or EAN)
  ean?: string;
  mainCategory?: string; // Main Category (e.g. Men)
  productType?: string; // Product Type (e.g. Bottom Wear)
  storeName: string;
  name: string; // Product Name
  toonLabel?: string; // Toon Label (e.g. AIM)
  color?: string; // Colour (e.g. LIGHT BLUE GREEN)
  size?: string; // Size (e.g. 28, 30, 32, 34, 36)
  columnExtra?: string; // "column" / Fabric (e.g. Cotton)
  styleNo?: string; // Style No (e.g. 5-P006-08)
  price?: number; // Price / MRP (e.g. 770)
  wondersoftQty: number; // Wondersoft Stock (calculated sum across ERP rows/sizes)
  ecommerceQty: number; // E-Com Stock
  targetQty: number; // Editable new stock quantity
  difference: number; // Diff (targetQty - ecommerceQty)
  status: StockStatus;
  hasDecreased: boolean;
  isOutOfStock: boolean;
  isInSync: boolean;
  isManuallyEdited?: boolean;
  matchedCount?: number; // Number of ERP rows summed
  matchedBreakdown?: string; // e.g. "Sizes summed: 28(5), 30(5), 32(6), 34(4), 36(5)"
  sourceWondersoftRow?: Record<string, any>;
  sourceWondersoftRows?: Record<string, any>[];
  sourceEcommerceRow?: Record<string, any>;
}

export interface ComparisonSummary {
  totalCompared: number;
  stockDecreasedCount: number;
  outOfStockCount: number;
  inSyncCount: number;
  stockIncreasedCount: number;
  onlyInWondersoftCount: number;
  onlyInEcommerceCount: number;
  totalUnitsDecreased: number;
  availableStores: string[];
  totalToonLabels?: number;
  distinctToonLabels?: string[];
}

export type TabKey = 'all' | 'changed' | 'decreased' | 'out_of_stock' | 'in_sync' | 'toon_summary';
