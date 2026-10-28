import React, { useState, useEffect } from 'react';
import { X, Check, ArrowRightLeft, Sparkles, ShieldCheck, Store, Tag, Hash } from 'lucide-react';
import { ColumnMapping, MatchMode, UploadedFile } from '../types';
import { extractUniqueStores } from '../utils/excel';

interface ColumnMappingModalProps {
  isOpen: boolean;
  onClose: () => void;
  wondersoftFile: UploadedFile;
  ecommerceFile: UploadedFile;
  initialMapping: ColumnMapping;
  onConfirmMapping: (mapping: ColumnMapping) => void;
}

export const ColumnMappingModal: React.FC<ColumnMappingModalProps> = ({
  isOpen,
  onClose,
  wondersoftFile,
  ecommerceFile,
  initialMapping,
  onConfirmMapping,
}) => {
  // Match Mode
  const [matchMode, setMatchMode] = useState<MatchMode>(initialMapping.matchMode || 'toon_label');

  // Toon Label mapping
  const [wondersoftToonLabelCol, setWondersoftToonLabelCol] = useState(
    initialMapping.wondersoftToonLabelColumn || initialMapping.toonLabelColumn || ''
  );
  const [ecommerceToonLabelCol, setEcommerceToonLabelCol] = useState(
    initialMapping.ecommerceToonLabelColumn || initialMapping.toonLabelColumn || ''
  );

  // Style No mapping
  const [wondersoftStyleNoCol, setWondersoftStyleNoCol] = useState(
    initialMapping.wondersoftStyleNoColumn || initialMapping.styleNoColumn || ''
  );
  const [ecommerceStyleNoCol, setEcommerceStyleNoCol] = useState(
    initialMapping.ecommerceStyleNoColumn || initialMapping.styleNoColumn || ''
  );

  // Optional EAN
  const [wondersoftSkuCol, setWondersoftSkuCol] = useState(
    initialMapping.wondersoftSkuColumn || initialMapping.skuColumn || ''
  );
  const [ecommerceSkuCol, setEcommerceSkuCol] = useState(
    initialMapping.ecommerceSkuColumn || initialMapping.skuColumn || ''
  );

  // Store Name state
  const [wondersoftStoreNameCol, setWondersoftStoreNameCol] = useState(
    initialMapping.wondersoftStoreNameColumn || ''
  );
  const [ecommerceStoreNameCol, setEcommerceStoreNameCol] = useState(
    initialMapping.ecommerceStoreNameColumn || initialMapping.storeNameColumn || ''
  );
  const [storeComparisonMode, setStoreComparisonMode] = useState<'all' | 'exact_store' | 'filter_store'>(
    initialMapping.storeComparisonMode || 'all'
  );
  const [selectedStore, setSelectedStore] = useState<string>(initialMapping.selectedStore || 'all');

  // Stock Quantities
  const [wondersoftQtyCol, setWondersoftQtyCol] = useState(
    initialMapping.wondersoftQtyColumn || ''
  );
  const [ecommerceQtyCol, setEcommerceQtyCol] = useState(
    initialMapping.ecommerceQtyColumn || ''
  );

  // Details
  const [ecommerceNameCol, setEcommerceNameCol] = useState(
    initialMapping.ecommerceNameColumn || initialMapping.nameColumn || ''
  );
  const [wondersoftNameCol, setWondersoftNameCol] = useState(
    initialMapping.wondersoftNameColumn || ''
  );
  const [mainCategoryCol, setMainCategoryCol] = useState(initialMapping.mainCategoryColumn || '');
  const [productTypeCol, setProductTypeCol] = useState(initialMapping.productTypeColumn || '');
  const [colorCol, setColorCol] = useState(
    initialMapping.ecommerceColorColumn || initialMapping.colorColumn || ''
  );
  const [wondersoftColorCol, setWondersoftColorCol] = useState(
    initialMapping.wondersoftColorColumn || ''
  );
  const [sizeCol, setSizeCol] = useState(
    initialMapping.ecommerceSizeColumn || initialMapping.sizeColumn || ''
  );
  const [wondersoftSizeCol, setWondersoftSizeCol] = useState(
    initialMapping.wondersoftSizeColumn || ''
  );
  const [columnExtraCol, setColumnExtraCol] = useState(initialMapping.columnExtraColumn || '');

  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  // Clean headers (filter out __EMPTY)
  const cleanWsHeaders = wondersoftFile.headers.filter((h) => h && !h.startsWith('__EMPTY'));
  const cleanEcomHeaders = ecommerceFile.headers.filter((h) => h && !h.startsWith('__EMPTY'));

  // Calculate available stores from files
  const availableStores = React.useMemo(() => {
    const ecomStores = extractUniqueStores(ecommerceFile.rawRows, ecommerceStoreNameCol);
    const wsStores = extractUniqueStores(wondersoftFile.rawRows, wondersoftStoreNameCol);
    const combined = Array.from(new Set([...ecomStores, ...wsStores])).filter(Boolean);
    return combined;
  }, [ecommerceFile.rawRows, wondersoftFile.rawRows, ecommerceStoreNameCol, wondersoftStoreNameCol]);

  // Synchronize initial mapping when modal opens or files change
  useEffect(() => {
    if (initialMapping.wondersoftToonLabelColumn) {
      setWondersoftToonLabelCol(initialMapping.wondersoftToonLabelColumn);
    }
    if (initialMapping.ecommerceToonLabelColumn) {
      setEcommerceToonLabelCol(initialMapping.ecommerceToonLabelColumn);
    }
    if (initialMapping.wondersoftStyleNoColumn) {
      setWondersoftStyleNoCol(initialMapping.wondersoftStyleNoColumn);
    }
    if (initialMapping.ecommerceStyleNoColumn) {
      setEcommerceStyleNoCol(initialMapping.ecommerceStyleNoColumn);
    }
    if (initialMapping.wondersoftQtyColumn) {
      setWondersoftQtyCol(initialMapping.wondersoftQtyColumn);
    }
    if (initialMapping.ecommerceQtyColumn) {
      setEcommerceQtyCol(initialMapping.ecommerceQtyColumn);
    }
  }, [initialMapping]);

  if (!isOpen) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    // Check that at least one match key is selected
    if (matchMode === 'toon_label' && !ecommerceToonLabelCol && !wondersoftToonLabelCol) {
      setErrorMsg('Please select the Toon Label column to calculate stock.');
      return;
    }
    if (matchMode === 'style_no' && !ecommerceStyleNoCol && !wondersoftStyleNoCol) {
      setErrorMsg('Please select the Style No column.');
      return;
    }
    if (matchMode === 'ean' && !ecommerceSkuCol && !wondersoftSkuCol) {
      setErrorMsg('Please select the EAN / Barcode column.');
      return;
    }

    if (!wondersoftQtyCol || !ecommerceQtyCol) {
      setErrorMsg('Please select the Stock Quantity column for both files.');
      return;
    }

    setErrorMsg(null);
    onConfirmMapping({
      matchMode,
      wondersoftToonLabelColumn: wondersoftToonLabelCol || undefined,
      ecommerceToonLabelColumn: ecommerceToonLabelCol || undefined,
      toonLabelColumn: ecommerceToonLabelCol || wondersoftToonLabelCol || undefined,
      wondersoftStyleNoColumn: wondersoftStyleNoCol || undefined,
      ecommerceStyleNoColumn: ecommerceStyleNoCol || undefined,
      styleNoColumn: ecommerceStyleNoCol || wondersoftStyleNoCol || undefined,
      wondersoftSkuColumn: wondersoftSkuCol || undefined,
      ecommerceSkuColumn: ecommerceSkuCol || undefined,
      skuColumn: ecommerceStyleNoCol || ecommerceSkuCol || wondersoftSkuCol || '',
      wondersoftNameColumn: wondersoftNameCol || undefined,
      ecommerceNameColumn: ecommerceNameCol || ecommerceStyleNoCol || '',
      nameColumn: ecommerceNameCol || wondersoftNameCol || '',
      wondersoftQtyColumn: wondersoftQtyCol,
      ecommerceQtyColumn: ecommerceQtyCol,
      mainCategoryColumn: mainCategoryCol || undefined,
      productTypeColumn: productTypeCol || undefined,
      wondersoftColorColumn: wondersoftColorCol || undefined,
      ecommerceColorColumn: colorCol || undefined,
      colorColumn: colorCol || wondersoftColorCol || undefined,
      wondersoftSizeColumn: wondersoftSizeCol || undefined,
      ecommerceSizeColumn: sizeCol || undefined,
      sizeColumn: sizeCol || wondersoftSizeCol || undefined,
      wondersoftStoreNameColumn: wondersoftStoreNameCol || undefined,
      ecommerceStoreNameColumn: ecommerceStoreNameCol || undefined,
      storeNameColumn: ecommerceStoreNameCol || wondersoftStoreNameCol || undefined,
      columnExtraColumn: columnExtraCol || undefined,
      storeComparisonMode,
      selectedStore: storeComparisonMode === 'filter_store' ? selectedStore : undefined,
    });
  };

  const previewWsRow = wondersoftFile.rawRows[0] || {};
  const previewEcomRow = ecommerceFile.rawRows[0] || {};

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div
        className="bg-white w-full max-w-2xl rounded-3xl overflow-hidden border border-slate-200 shadow-2xl animate-in zoom-in-95 duration-150 flex flex-col max-h-[92vh]"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Modal Header */}
        <div className="p-5 sm:p-6 border-b border-slate-100 flex items-center justify-between bg-blue-50/70">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-xs">
              <ArrowRightLeft className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                Column Mapping
                <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-blue-100 text-blue-800">
                  <Sparkles className="w-3 h-3 inline mr-1" />
                  Toon Label Setup
                </span>
              </h3>
              <p className="text-xs text-slate-500">
                Match by Toon Label and sum ERP closing stock across sizes
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2 rounded-xl text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Guarantee Banner */}
        <div className="bg-emerald-50 px-5 sm:px-6 py-2.5 border-b border-emerald-100 flex items-center gap-2 text-xs font-semibold text-emerald-800">
          <ShieldCheck className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>
            Accurate Stock Calculation: Sums all size rows in Wondersoft ERP to compare with your E-Commerce stock.
          </span>
        </div>

        {/* Mapping Form */}
        <form onSubmit={handleSubmit} className="p-5 sm:p-6 overflow-y-auto space-y-4">
          {errorMsg && (
            <div className="p-3 rounded-xl bg-rose-50 border border-rose-200 text-rose-700 text-xs font-medium">
              {errorMsg}
            </div>
          )}

          {/* 1. Comparison Match Key: Toon Label */}
          <div className="bg-blue-50/50 p-4 rounded-2xl border border-blue-200 space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-sm font-bold text-blue-950 flex items-center gap-2">
                <div className="w-5 h-5 rounded-full bg-blue-600 text-white text-xs flex items-center justify-center font-bold">
                  1
                </div>
                <Tag className="w-4 h-4 text-blue-600" />
                Primary Matching Key: Toon Label
              </label>
              <span className="text-[11px] font-bold text-blue-700 bg-blue-100 px-2 py-0.5 rounded-full">
                Recommended
              </span>
            </div>

            {/* Match Mode Selector */}
            <div className="grid grid-cols-3 gap-2 pt-1">
              <button
                type="button"
                onClick={() => setMatchMode('toon_label')}
                className={`py-2 px-2.5 rounded-xl text-xs font-semibold border text-center transition-all cursor-pointer ${
                  matchMode === 'toon_label'
                    ? 'bg-blue-600 text-white border-blue-600 shadow-xs'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                <div className="font-bold">Toon Label</div>
                <div className={`text-[10px] ${matchMode === 'toon_label' ? 'text-blue-100' : 'text-slate-400'}`}>
                  Toon Label + Style
                </div>
              </button>

              <button
                type="button"
                onClick={() => setMatchMode('style_no')}
                className={`py-2 px-2.5 rounded-xl text-xs font-semibold border text-center transition-all cursor-pointer ${
                  matchMode === 'style_no'
                    ? 'bg-blue-600 text-white border-blue-600 shadow-xs'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                <div className="font-bold">Style No</div>
                <div className={`text-[10px] ${matchMode === 'style_no' ? 'text-blue-100' : 'text-slate-400'}`}>
                  By Style Code
                </div>
              </button>

              <button
                type="button"
                onClick={() => setMatchMode('ean')}
                className={`py-2 px-2.5 rounded-xl text-xs font-semibold border text-center transition-all cursor-pointer ${
                  matchMode === 'ean'
                    ? 'bg-blue-600 text-white border-blue-600 shadow-xs'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                <div className="font-bold">EAN / Barcode</div>
                <div className={`text-[10px] ${matchMode === 'ean' ? 'text-blue-100' : 'text-slate-400'}`}>
                  If numbers exist
                </div>
              </button>
            </div>

            {/* Toon Label Dropdowns */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  Wondersoft ERP (Toon Label):
                </span>
                <select
                  value={wondersoftToonLabelCol}
                  onChange={(e) => setWondersoftToonLabelCol(e.target.value)}
                  className="w-full text-xs font-bold py-2 px-3 rounded-xl bg-white border border-blue-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-xs"
                >
                  <option value="">-- Select Column --</option>
                  {cleanWsHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
                <div className="text-[11px] text-slate-500 mt-1 truncate">
                  Sample: <span className="font-bold text-blue-700">{String(previewWsRow[wondersoftToonLabelCol] || '—')}</span>
                </div>
              </div>

              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  E-Commerce Store (Toon Label):
                </span>
                <select
                  value={ecommerceToonLabelCol}
                  onChange={(e) => setEcommerceToonLabelCol(e.target.value)}
                  className="w-full text-xs font-bold py-2 px-3 rounded-xl bg-white border border-blue-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-xs"
                >
                  <option value="">-- Select Column --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
                <div className="text-[11px] text-slate-500 mt-1 truncate">
                  Sample: <span className="font-bold text-blue-700">{String(previewEcomRow[ecommerceToonLabelCol] || '—')}</span>
                </div>
              </div>
            </div>

            {/* Style No Dropdowns */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1 border-t border-blue-200/60">
              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  Wondersoft Style No / Article No:
                </span>
                <select
                  value={wondersoftStyleNoCol}
                  onChange={(e) => setWondersoftStyleNoCol(e.target.value)}
                  className="w-full text-xs font-medium py-2 px-3 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- Select Column --</option>
                  {cleanWsHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
                <div className="text-[11px] text-slate-500 mt-1 truncate">
                  Sample: <span className="font-mono text-slate-800">{String(previewWsRow[wondersoftStyleNoCol] || '—')}</span>
                </div>
              </div>

              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  E-Commerce Style No (e.g. 5-P006-08):
                </span>
                <select
                  value={ecommerceStyleNoCol}
                  onChange={(e) => setEcommerceStyleNoCol(e.target.value)}
                  className="w-full text-xs font-medium py-2 px-3 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- Select Column --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
                <div className="text-[11px] text-slate-500 mt-1 truncate">
                  Sample: <span className="font-mono text-slate-800">{String(previewEcomRow[ecommerceStyleNoCol] || '—')}</span>
                </div>
              </div>
            </div>
          </div>

          {/* 2. Stock Quantity Columns */}
          <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-sm font-bold text-slate-900 flex items-center gap-2">
                <div className="w-5 h-5 rounded-full bg-blue-600 text-white text-xs flex items-center justify-center font-bold">
                  2
                </div>
                <Hash className="w-4 h-4 text-blue-600" />
                Stock Quantity Columns (Compare Values)
              </label>
              <span className="text-[11px] font-semibold text-blue-600">Required</span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  Wondersoft ERP Stock (Closing Qty):
                </span>
                <select
                  value={wondersoftQtyCol}
                  onChange={(e) => setWondersoftQtyCol(e.target.value)}
                  className="w-full text-xs font-semibold py-2 px-3 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  {cleanWsHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
                <div className="text-[11px] text-slate-500 mt-1 truncate">
                  Sample Stock: <span className="font-mono font-bold text-blue-600">{String(previewWsRow[wondersoftQtyCol] ?? 0)}</span>
                </div>
              </div>

              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  E-Commerce Store Stock:
                </span>
                <select
                  value={ecommerceQtyCol}
                  onChange={(e) => setEcommerceQtyCol(e.target.value)}
                  className="w-full text-xs font-semibold py-2 px-3 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
                <div className="text-[11px] text-slate-500 mt-1 truncate">
                  Sample Stock: <span className="font-mono font-bold text-slate-800">{String(previewEcomRow[ecommerceQtyCol] ?? 0)}</span>
                </div>
              </div>
            </div>
          </div>

          {/* 3. Separate Store Name Comparison */}
          <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-sm font-bold text-slate-900 flex items-center gap-2">
                <div className="w-5 h-5 rounded-full bg-slate-800 text-white text-xs flex items-center justify-center font-bold">
                  3
                </div>
                <Store className="w-4 h-4 text-slate-700" />
                Store Name (Optional Store Filter)
              </label>
              <span className="text-[11px] font-semibold text-slate-500">Store Filter</span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  Wondersoft Store Name:
                </span>
                <select
                  value={wondersoftStoreNameCol}
                  onChange={(e) => setWondersoftStoreNameCol(e.target.value)}
                  className="w-full text-xs font-semibold py-2 px-3 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- No Store Column --</option>
                  {cleanWsHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">
                  E-Commerce Store Name:
                </span>
                <select
                  value={ecommerceStoreNameCol}
                  onChange={(e) => setEcommerceStoreNameCol(e.target.value)}
                  className="w-full text-xs font-semibold py-2 px-3 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- No Store Column --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Store Comparison Mode Options */}
            <div className="pt-2 border-t border-slate-200/80">
              <span className="text-xs font-semibold text-slate-700 mb-2 block">
                Store Comparison Method:
              </span>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                <button
                  type="button"
                  onClick={() => setStoreComparisonMode('all')}
                  className={`px-3 py-2 rounded-xl text-xs font-semibold text-left border transition-all cursor-pointer ${
                    storeComparisonMode === 'all'
                      ? 'bg-blue-600 text-white border-blue-600 shadow-xs'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  <div className="font-bold">All Stores</div>
                  <div className={`text-[10px] ${storeComparisonMode === 'all' ? 'text-blue-100' : 'text-slate-500'}`}>
                    Compare across entire file
                  </div>
                </button>

                <button
                  type="button"
                  onClick={() => setStoreComparisonMode('exact_store')}
                  className={`px-3 py-2 rounded-xl text-xs font-semibold text-left border transition-all cursor-pointer ${
                    storeComparisonMode === 'exact_store'
                      ? 'bg-blue-600 text-white border-blue-600 shadow-xs'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  <div className="font-bold">Match Same Store</div>
                  <div className={`text-[10px] ${storeComparisonMode === 'exact_store' ? 'text-blue-100' : 'text-slate-500'}`}>
                    Match Store + Toon Label
                  </div>
                </button>

                <button
                  type="button"
                  onClick={() => setStoreComparisonMode('filter_store')}
                  className={`px-3 py-2 rounded-xl text-xs font-semibold text-left border transition-all cursor-pointer ${
                    storeComparisonMode === 'filter_store'
                      ? 'bg-blue-600 text-white border-blue-600 shadow-xs'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  <div className="font-bold">Filter By Store</div>
                  <div className={`text-[10px] ${storeComparisonMode === 'filter_store' ? 'text-blue-100' : 'text-slate-500'}`}>
                    Select 1 store
                  </div>
                </button>
              </div>

              {storeComparisonMode === 'filter_store' && availableStores.length > 0 && (
                <div className="mt-3 p-3 bg-white rounded-xl border border-blue-200">
                  <label className="text-xs font-bold text-slate-800 block mb-1">
                    Select Store:
                  </label>
                  <select
                    value={selectedStore}
                    onChange={(e) => setSelectedStore(e.target.value)}
                    className="w-full text-xs font-semibold py-2 px-3 rounded-lg border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="all">All Stores</option>
                    {availableStores.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>
          </div>

          {/* 4. Product Details */}
          <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 space-y-3">
            <label className="text-sm font-bold text-slate-900 flex items-center gap-2">
              <div className="w-5 h-5 rounded-full bg-slate-800 text-white text-xs flex items-center justify-center font-bold">
                4
              </div>
              Product Details (Product Name, Colour, Size, Category)
            </label>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">Product Name:</span>
                <select
                  value={ecommerceNameCol}
                  onChange={(e) => setEcommerceNameCol(e.target.value)}
                  className="w-full text-xs font-medium py-2 px-2.5 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- None --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">Colour:</span>
                <select
                  value={colorCol}
                  onChange={(e) => setColorCol(e.target.value)}
                  className="w-full text-xs font-medium py-2 px-2.5 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- None --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">Size:</span>
                <select
                  value={sizeCol}
                  onChange={(e) => setSizeCol(e.target.value)}
                  className="w-full text-xs font-medium py-2 px-2.5 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- None --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">Main Category / Dept:</span>
                <select
                  value={mainCategoryCol}
                  onChange={(e) => setMainCategoryCol(e.target.value)}
                  className="w-full text-xs font-medium py-2 px-2.5 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- None --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>



              <div>
                <span className="text-xs font-semibold text-slate-700 mb-1 block">Column / Fabric:</span>
                <select
                  value={columnExtraCol}
                  onChange={(e) => setColumnExtraCol(e.target.value)}
                  className="w-full text-xs font-medium py-2 px-2.5 rounded-xl bg-white border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">-- None --</option>
                  {cleanEcomHeaders.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Footer Buttons */}
          <div className="pt-2 flex items-center justify-end gap-3 border-t border-slate-100">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 text-xs font-semibold rounded-xl text-slate-600 hover:bg-slate-100 transition-colors cursor-pointer"
            >
              Cancel
            </button>
            <button
              type="submit"
              id="confirm-mapping-submit-btn"
              className="px-6 py-2.5 text-xs font-bold rounded-xl bg-blue-600 hover:bg-blue-700 text-white shadow-xs transition-all cursor-pointer flex items-center gap-2 active:scale-95"
            >
              <Check className="w-4 h-4" />
              <span>Confirm & Compare Stocks</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
