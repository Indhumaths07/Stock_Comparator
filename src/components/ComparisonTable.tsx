import React, { useState, useMemo } from 'react';
import {
  Search,
  Download,
  Copy,
  ChevronLeft,
  ChevronRight,
  Store,
  Tag,
  Layers,
  BarChart3,
  List,
} from 'lucide-react';
import { ColumnMapping, ComparisonItem, TabKey, UploadedFile } from '../types';
import { aggregateByToonLabel, exportChangedStockReport } from '../utils/excel';

interface ComparisonTableProps {
  items: ComparisonItem[];
  activeTab?: TabKey;
  onTabChange?: (tab: TabKey) => void;
  onCopyNotice: (msg: string) => void;
  onUpdateTargetQty?: (itemId: string, newQty: number) => void;
  ecommerceFile?: UploadedFile | null;
  columnMapping?: ColumnMapping;
}

export const ComparisonTable: React.FC<ComparisonTableProps> = ({
  items,
  onCopyNotice,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedStoreFilter, setSelectedStoreFilter] = useState<string>('all');
  const [selectedToonFilter, setSelectedToonFilter] = useState<string>('all');
  const [viewMode, setViewMode] = useState<'products' | 'toon_summary'>('products');
  const [sortField, setSortField] = useState<
    | 'mainCategory'
    | 'name'
    | 'toonLabel'
    | 'styleNo'
    | 'wondersoftQty'
    | 'ecommerceQty'
    | 'difference'
  >('difference');
  const [sortAsc, setSortAsc] = useState(true);
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 25;

  // Distinct store list for store filtering
  const availableStores = useMemo(() => {
    const set = new Set<string>();
    for (const item of items) {
      if (item.storeName) set.add(item.storeName);
    }
    return Array.from(set).sort();
  }, [items]);

  // Distinct Toon Labels for quick filtering
  const availableToonLabels = useMemo(() => {
    const set = new Set<string>();
    for (const item of items) {
      if (item.toonLabel) set.add(item.toonLabel);
    }
    return Array.from(set).sort();
  }, [items]);

  // Toon Label aggregated summary
  const toonSummaryList = useMemo(() => {
    return aggregateByToonLabel(items);
  }, [items]);

  // Filter items based on search, store, and toon label
  const filteredItems = useMemo(() => {
    return items.filter((item) => {
      // Store filter
      if (selectedStoreFilter !== 'all') {
        if (item.storeName.toLowerCase() !== selectedStoreFilter.toLowerCase()) return false;
      }

      // Toon Label filter
      if (selectedToonFilter !== 'all') {
        if ((item.toonLabel || '').toLowerCase() !== selectedToonFilter.toLowerCase()) return false;
      }

      // Search filter (Item Name, Toon Label, Style No, Color, Category)
      if (searchTerm.trim() !== '') {
        const query = searchTerm.toLowerCase();
        const toonMatch = (item.toonLabel || '').toLowerCase().includes(query);
        const styleMatch = (item.styleNo || item.sku || '').toLowerCase().includes(query);
        const nameMatch = item.name.toLowerCase().includes(query);
        const colorMatch = (item.color || '').toLowerCase().includes(query);
        const categoryMatch = item.mainCategory ? item.mainCategory.toLowerCase().includes(query) : false;
        const storeMatch = item.storeName.toLowerCase().includes(query);
        return toonMatch || styleMatch || nameMatch || colorMatch || categoryMatch || storeMatch;
      }

      return true;
    });
  }, [items, searchTerm, selectedStoreFilter, selectedToonFilter]);

  // Sort filtered items
  const sortedItems = useMemo(() => {
    return [...filteredItems].sort((a, b) => {
      let valA: any = a[sortField];
      let valB: any = b[sortField];

      if (typeof valA === 'string') {
        return sortAsc
          ? (valA || '').localeCompare(valB || '')
          : (valB || '').localeCompare(valA || '');
      }

      return sortAsc ? (valA ?? 0) - (valB ?? 0) : (valB ?? 0) - (valA ?? 0);
    });
  }, [filteredItems, sortField, sortAsc]);

  // Pagination
  const totalPages = Math.ceil(sortedItems.length / pageSize) || 1;
  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return sortedItems.slice(start, start + pageSize);
  }, [sortedItems, currentPage, pageSize]);

  const handleSort = (
    field:
      | 'mainCategory'
      | 'name'
      | 'toonLabel'
      | 'styleNo'
      | 'wondersoftQty'
      | 'ecommerceQty'
      | 'difference'
  ) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const handleCopySingleText = (text: string, label: string) => {
    navigator.clipboard.writeText(text);
    onCopyNotice(`Copied ${label}: ${text}`);
  };

  const handleCopyAllToonLabels = () => {
    const labels = Array.from(new Set(sortedItems.map((i) => i.toonLabel || i.name))).join('\n');
    navigator.clipboard.writeText(labels);
    onCopyNotice(`Copied Toon Labels to clipboard`);
  };

  // Direct Excel Download with rich styling
  const handleDownloadExcel = () => {
    const { rowCount, filename } = exportChangedStockReport(
      items,
      undefined,
      undefined,
      { onlyChanged: false }
    );
    onCopyNotice(`✓ Downloaded ${filename} (${rowCount} rows). In-Sync (Diff 0) highlighted in green!`);
  };

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
      {/* Top Bar: Title, Count, View Mode Switch, and Excel Download */}
      <div className="p-4 sm:p-5 border-b border-slate-200 bg-slate-50/60 flex flex-col gap-4">
        {/* Row 1: Header title + View Switcher + Download Excel button */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center shrink-0">
              <Layers className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                <span>Stock Comparison</span>
                <span className="px-2 py-0.5 rounded-full text-xs font-bold bg-blue-100 text-blue-800">
                  {items.length.toLocaleString()} items
                </span>
                {availableToonLabels.length > 0 && (
                  <span className="px-2 py-0.5 rounded-full text-xs font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200">
                    <Tag className="w-3 h-3 inline mr-1" />
                    {availableToonLabels.length} Toon Labels
                  </span>
                )}
              </h3>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap sm:flex-nowrap">
            {/* View Toggle: Product List vs Toon Label Grouped Summary */}
            <div className="inline-flex rounded-xl bg-slate-200/80 p-0.5 text-xs font-semibold">
              <button
                onClick={() => setViewMode('products')}
                className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                  viewMode === 'products'
                    ? 'bg-white text-blue-700 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <List className="w-3.5 h-3.5" />
                <span>Products List</span>
              </button>
              <button
                onClick={() => setViewMode('toon_summary')}
                className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                  viewMode === 'toon_summary'
                    ? 'bg-white text-blue-700 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <BarChart3 className="w-3.5 h-3.5" />
                <span>Toon Label Summary</span>
              </button>
            </div>

            {/* Single Prominent Excel (.xlsx) Download */}
            <button
              onClick={handleDownloadExcel}
              id="download-excel-btn"
              title="Download Excel file with stock comparison results"
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-xl bg-blue-600 hover:bg-blue-700 text-white transition-all shadow-xs cursor-pointer active:scale-95"
            >
              <Download className="w-4 h-4" />
              <span>Download Excel (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* Row 2: Search Bar + Toon Label Filter + Store Name Filter */}
        <div className="flex flex-col sm:flex-row items-center gap-3">
          <div className="relative flex-1 w-full">
            <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1);
              }}
              placeholder="Search by Toon Label, Style No (e.g. 5-P006-08), Product Name, or Colour..."
              className="w-full pl-10 pr-4 py-2 text-xs sm:text-sm font-medium rounded-xl bg-white border border-slate-300 text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-inner"
            />
          </div>

          {/* Dedicated Toon Label Filter Dropdown */}
          {availableToonLabels.length > 0 && (
            <div className="flex items-center gap-2 w-full sm:w-auto shrink-0 bg-white px-3 py-1.5 rounded-xl border border-slate-300 shadow-xs">
              <Tag className="w-4 h-4 text-blue-600 shrink-0" />
              <span className="text-xs font-bold text-slate-700 whitespace-nowrap">
                Toon Label:
              </span>
              <select
                value={selectedToonFilter}
                onChange={(e) => {
                  setSelectedToonFilter(e.target.value);
                  setCurrentPage(1);
                }}
                className="text-xs font-semibold bg-transparent text-slate-900 focus:outline-none cursor-pointer max-w-[130px] truncate"
              >
                <option value="all">All Labels ({availableToonLabels.length})</option>
                {availableToonLabels.map((lbl) => (
                  <option key={lbl} value={lbl}>
                    {lbl}
                  </option>
                ))}
              </select>
            </div>
          )}

          {/* Dedicated Store Name Filter Dropdown */}
          {availableStores.length > 0 && (
            <div className="flex items-center gap-2 w-full sm:w-auto shrink-0 bg-white px-3 py-1.5 rounded-xl border border-slate-300 shadow-xs">
              <Store className="w-4 h-4 text-slate-600 shrink-0" />
              <span className="text-xs font-bold text-slate-700 whitespace-nowrap">
                Store:
              </span>
              <select
                value={selectedStoreFilter}
                onChange={(e) => {
                  setSelectedStoreFilter(e.target.value);
                  setCurrentPage(1);
                }}
                className="text-xs font-semibold bg-transparent text-slate-900 focus:outline-none cursor-pointer max-w-[120px] truncate"
              >
                <option value="all">All Stores</option>
                {availableStores.map((store) => (
                  <option key={store} value={store}>
                    {store}
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>
      </div>

      {/* VIEW 1: PRODUCTS DETAILED LIST */}
      {viewMode === 'products' ? (
        <>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-blue-50/60 text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                  <th
                    onClick={() => handleSort('mainCategory')}
                    className="py-3 px-3 cursor-pointer hover:text-blue-600 transition-colors whitespace-nowrap"
                  >
                    Main Category
                  </th>

                  <th
                    onClick={() => handleSort('name')}
                    className="py-3 px-3 cursor-pointer hover:text-blue-600 transition-colors min-w-[180px]"
                  >
                    Product Name
                  </th>

                  <th
                    onClick={() => handleSort('toonLabel')}
                    className="py-3 px-3 cursor-pointer hover:text-blue-600 transition-colors whitespace-nowrap"
                  >
                    <div className="flex items-center gap-1">
                      <Tag className="w-3.5 h-3.5 text-blue-600" />
                      <span>Toon Label</span>
                    </div>
                  </th>

                  <th
                    onClick={() => handleSort('styleNo')}
                    className="py-3 px-3 cursor-pointer hover:text-blue-600 transition-colors whitespace-nowrap"
                  >
                    Style No
                  </th>

                  <th className="py-3 px-3 whitespace-nowrap">
                    Colour / Size
                  </th>

                  <th
                    onClick={() => handleSort('ecommerceQty')}
                    className="py-3 px-3 text-right cursor-pointer hover:text-blue-600 transition-colors whitespace-nowrap"
                  >
                    E-Com Stock
                  </th>

                  <th
                    onClick={() => handleSort('wondersoftQty')}
                    className="py-3 px-3 text-right cursor-pointer hover:text-blue-600 transition-colors whitespace-nowrap"
                  >
                    Wondersoft Stock
                  </th>

                  <th
                    onClick={() => handleSort('difference')}
                    className="py-3 px-3 text-right cursor-pointer hover:text-blue-600 transition-colors whitespace-nowrap"
                  >
                    Diff
                  </th>
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100">
                {paginatedItems.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="py-12 text-center text-slate-500">
                      <div className="flex flex-col items-center justify-center gap-2">
                        <p className="font-semibold text-slate-700">No matching products found</p>
                        <p className="text-xs text-slate-400">
                          Try adjusting search term or Toon Label filter.
                        </p>
                      </div>
                    </td>
                  </tr>
                ) : (
                  paginatedItems.map((item) => {
                    const isDecreased = item.hasDecreased;
                    const isOos = item.isOutOfStock;
                    const isZeroDiff = item.difference === 0;

                    return (
                      <tr
                        key={item.id}
                        className={`hover:bg-blue-50/40 transition-colors ${
                          isOos
                            ? 'bg-rose-50/30'
                            : isDecreased
                            ? 'bg-amber-50/25'
                            : isZeroDiff
                            ? 'bg-emerald-50/30'
                            : ''
                        }`}
                      >
                        {/* 1. Main Category */}
                        <td className="py-3 px-3 font-medium text-slate-800 whitespace-nowrap">
                          {item.mainCategory ? (
                            <span className="px-2 py-0.5 rounded-md bg-slate-100 text-slate-800 text-[11px] font-semibold border border-slate-200">
                              {item.mainCategory}
                            </span>
                          ) : (
                            <span className="text-slate-400">—</span>
                          )}
                        </td>

                        {/* 2. Product Name */}
                        <td className="py-3 px-3 text-slate-800 font-medium max-w-[210px]">
                          <div className="truncate font-semibold text-slate-900" title={item.name}>
                            {item.name}
                          </div>
                          {item.columnExtra && (
                            <div className="text-[10px] text-slate-400 truncate">
                              {item.columnExtra}
                            </div>
                          )}
                        </td>

                        {/* 4. Toon Label */}
                        <td className="py-3 px-3 whitespace-nowrap">
                          {item.toonLabel ? (
                            <div className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-blue-50 text-blue-800 font-bold text-xs border border-blue-200">
                              <span>{item.toonLabel}</span>
                              <button
                                onClick={() => handleCopySingleText(item.toonLabel || '', 'Toon Label')}
                                className="text-blue-400 hover:text-blue-800 p-0.5 rounded cursor-pointer"
                                title="Copy Toon Label"
                              >
                                <Copy className="w-2.5 h-2.5" />
                              </button>
                            </div>
                          ) : (
                            <span className="text-slate-400">—</span>
                          )}
                        </td>

                        {/* 5. Style No */}
                        <td className="py-3 px-3 whitespace-nowrap">
                          <span className="font-mono font-bold text-slate-800">
                            {item.styleNo || item.sku || '—'}
                          </span>
                        </td>

                        {/* 6. Colour / Size */}
                        <td className="py-3 px-3 max-w-[160px]">
                          <div className="truncate text-slate-700 font-medium text-[11px]">
                            {item.color || '—'}
                          </div>
                          {item.size && (
                            <div className="truncate text-[10px] text-slate-400 font-mono">
                              Size: {item.size}
                            </div>
                          )}
                        </td>

                        {/* 7. Current E-Com Stock */}
                        <td className="py-3 px-3 text-right font-mono font-medium text-slate-600">
                          {item.ecommerceQty.toLocaleString()}
                        </td>

                        {/* 8. Wondersoft ERP Stock (Calculated) */}
                        <td className="py-3 px-3 text-right">
                          <div className="font-mono font-bold text-blue-700 text-xs">
                            {item.wondersoftQty.toLocaleString()}
                          </div>
                          {item.matchedBreakdown && (
                            <div
                              className="text-[10px] text-slate-400 truncate font-sans max-w-[140px] ml-auto cursor-help"
                              title={item.matchedBreakdown}
                            >
                              {item.matchedBreakdown}
                            </div>
                          )}
                        </td>

                        {/* 9. Difference (Absolute number) */}
                        <td className="py-3 px-3 text-right font-mono font-bold">
                          {isZeroDiff ? (
                            <span className="inline-flex items-center justify-center px-2 py-0.5 rounded-md bg-emerald-100 text-emerald-800 text-xs font-bold border border-emerald-200">
                              0
                            </span>
                          ) : (
                            <span className="text-slate-800">{Math.abs(item.difference)}</span>
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Table Pagination & Batch Copy Footer */}
          <div className="p-4 border-t border-slate-200 bg-slate-50/50 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-2">
              <button
                onClick={handleCopyAllToonLabels}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 font-semibold transition-colors cursor-pointer"
              >
                <Copy className="w-3.5 h-3.5 text-slate-500" />
                <span>Copy Toon Labels</span>
              </button>
            </div>

            <div className="flex items-center gap-3">
              <span className="text-slate-500">
                Showing{' '}
                <span className="font-semibold text-slate-800">
                  {sortedItems.length === 0 ? 0 : (currentPage - 1) * pageSize + 1}
                </span>{' '}
                to{' '}
                <span className="font-semibold text-slate-800">
                  {Math.min(currentPage * pageSize, sortedItems.length)}
                </span>{' '}
                of <span className="font-semibold text-slate-800">{sortedItems.length}</span>
              </span>

              <div className="flex items-center gap-1">
                <button
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  disabled={currentPage <= 1}
                  className="p-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <span className="px-2 font-bold text-slate-800">
                  {currentPage} / {totalPages}
                </span>
                <button
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  disabled={currentPage >= totalPages}
                  className="p-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          </div>
        </>
      ) : (
        /* VIEW 2: TOON LABEL SUMMARY GROUPED VIEW */
        <div className="p-4 sm:p-5">
          <div className="mb-4">
            <h4 className="text-sm font-bold text-slate-900">
              Toon Label Stock Calculation Summary
            </h4>
            <p className="text-xs text-slate-500">
              Aggregated stock totals grouped by Toon Label across all sizes and styles.
            </p>
          </div>

          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-blue-50/60 text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                  <th className="py-3 px-3">Toon Label</th>
                  <th className="py-3 px-3 text-center">Styles Count</th>
                  <th className="py-3 px-3 text-right">Total E-Com Stock</th>
                  <th className="py-3 px-3 text-right">Total Wondersoft Stock</th>
                  <th className="py-3 px-3 text-right">Net Difference</th>
                  <th className="py-3 px-3 text-center">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {toonSummaryList.map((row) => {
                  const isSync = row.netDifference === 0;
                  return (
                    <tr
                      key={row.toonLabel}
                      className={`hover:bg-blue-50/30 transition-colors ${
                        isSync ? 'bg-emerald-50/20' : 'bg-amber-50/20'
                      }`}
                    >
                      <td className="py-3 px-3">
                        <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-blue-100 text-blue-900 font-bold text-xs border border-blue-200">
                          <Tag className="w-3 h-3 text-blue-600" />
                          {row.toonLabel}
                        </span>
                        {row.styles.length > 0 && (
                          <div className="text-[10px] text-slate-400 mt-0.5 truncate max-w-[200px]">
                            {row.styles.join(', ')}
                          </div>
                        )}
                      </td>

                      <td className="py-3 px-3 text-center font-semibold text-slate-700">
                        {row.itemCount} styles
                      </td>

                      <td className="py-3 px-3 text-right font-mono font-medium text-slate-700">
                        {row.totalEcommerceQty.toLocaleString()}
                      </td>

                      <td className="py-3 px-3 text-right font-mono font-bold text-blue-700">
                        {row.totalWondersoftQty.toLocaleString()}
                      </td>

                      <td className="py-3 px-3 text-right font-mono font-bold">
                        {row.netDifference === 0 ? (
                          <span className="inline-flex items-center justify-center px-2 py-0.5 rounded-md bg-emerald-100 text-emerald-800 text-xs font-bold border border-emerald-200">
                            0
                          </span>
                        ) : (
                          <span className="text-slate-800">{Math.abs(row.netDifference)}</span>
                        )}
                      </td>

                      <td className="py-3 px-3 text-center">
                        {row.netDifference === 0 ? (
                          <span className="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[11px] font-bold">
                            In Sync
                          </span>
                        ) : (
                          <span className="px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 text-[11px] font-bold">
                            {row.decreasedCount} reduced
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
