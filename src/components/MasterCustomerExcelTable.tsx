import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  Eye,
  PenLine,
  Printer,
  Save,
  Filter,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Check,
  X,
  Search,
  RotateCcw,
  Download,
  Trash2,
  Phone,
  Maximize2,
  Minimize2,
  Table as TableIcon,
  Columns
} from "lucide-react";
import * as XLSX from "xlsx";

export interface MasterCustomerExcelTableProps {
  customers: any[];
  language?: "te" | "en";
  onSave: (chassisNo: string, updatedFields: any) => void;
  onDelete?: (chassisNo: string) => void;
  onView?: (customer: any) => void;
}

export const MasterCustomerExcelTable: React.FC<MasterCustomerExcelTableProps> = ({
  customers,
  language = "te",
  onSave,
  onDelete,
  onView
}) => {
  const isTe = language === "te";

  const [viewMode, setViewMode] = useState<"5col" | "detailed">("5col");
  const [rowDensity, setRowDensity] = useState<"compact" | "normal" | "spacious">("normal");

  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(25);

  const [sortCol, setSortCol] = useState<string>("Customer Name");
  const [sortDir, setSortDir] = useState<"desc" | "asc">("asc");

  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [filterSearchText, setFilterSearchText] = useState<string>("");

  const [savingRows, setSavingRows] = useState<Record<string, boolean>>({});
  const [savedSuccessRows, setSavedSuccessRows] = useState<Record<string, boolean>>({});
  const [rowDrafts, setRowDrafts] = useState<Record<string, any>>({});

  const getRowKey = (cust: any, globalIdx: number) => {
    const chassis = String(cust["Chassis no"] || cust.chassisNo || cust.chassis || cust.__chassisDisplay || "").trim();
    if (chassis) return `chassis_${chassis}`;
    return `row_${globalIdx}_${String(cust["Customer Name"] || cust.custName || "").trim()}`;
  };

  const getCustValue = (cust: any, colKey: string, globalIdx: number = 0) => {
    const key = getRowKey(cust, globalIdx);
    if (rowDrafts[key] && rowDrafts[key][colKey] !== undefined) {
      return rowDrafts[key][colKey];
    }
    return getColDisplayValue(cust, colKey);
  };

  const handleFieldChange = (cust: any, globalIdx: number, colKey: string, value: any) => {
    const key = getRowKey(cust, globalIdx);
    setRowDrafts((prev) => ({
      ...prev,
      [key]: {
        ...(prev[key] || {}),
        [colKey]: value,
      },
    }));
  };

  const getColDisplayValue = (cust: any, colKey: string): string => {
    const fd = cust.fullData || {};
    switch (colKey) {
      case "SUPERVISOR":
        return String(cust["SUPERVISOR"] || cust.supervisor || cust.wsIncharge || fd["SUPERVISOR"] || fd.supervisor || fd.wsIncharge || "").trim();
      case "BRANCH":
      case "Branch":
        return String(cust["BRANCH"] || cust["Branch"] || cust.branch || fd["BRANCH"] || fd["Branch"] || fd.branch || "").trim();
      case "SL.NO":
        return String(cust["SL.NO"] || cust["SL. No"] || cust["S.No"] || cust["S.No."] || cust.slNo || cust.sNo || fd["SL.NO"] || fd["SL. No"] || fd["S.No"] || fd.slNo || "").trim();
      case "Model":
      case "model":
        return String(cust["Model"] || cust.model || fd["Model"] || fd.model || "").trim();
      case "MODEL TYPE":
      case "modelType":
        return String(cust["MODEL TYPE"] || cust.modelType || fd["MODEL TYPE"] || fd.modelType || "").trim();
      case "Chassis no":
      case "chassisNo":
        return String(cust["Chassis no"] || cust.chassisNo || cust.chassis || fd["Chassis no"] || fd.chassisNo || fd.chassis || "").trim();
      case "Engine No:":
      case "Engine No":
        return String(cust["Engine No:"] || cust["Engine No"] || cust.engineNo || fd["Engine No:"] || fd["Engine No"] || fd.engineNo || "").trim();
      case "Date of del":
      case "Date of Delivery":
        return String(cust["Date of del"] || cust["Date of Delivery"] || cust["DEL DATE"] || cust.dateOfDelivery || cust.installDate || fd["Date of del"] || fd["Date of Delivery"] || fd.dateOfDelivery || "").trim();
      case "Customer Name":
      case "custName":
        return String(cust["Customer Name"] || cust.custName || cust.customerName || fd["Customer Name"] || fd.custName || fd.customerName || "").trim();
      case "FATHER NAME":
      case "Father":
        return String(cust["FATHER NAME"] || cust["Father"] || cust.fatherName || fd["FATHER NAME"] || fd["Father"] || fd.fatherName || "").trim();
      case "ADDRESS":
      case "address":
        return String(cust["ADDRESS"] || cust["Address"] || cust.address || cust.custAddr || fd["ADDRESS"] || fd["Address"] || fd.address || fd.custAddr || "").trim();
      case "VILLAGE":
      case "Village":
        return String(cust["VILLAGE"] || cust["Village"] || cust.village || fd["VILLAGE"] || fd["Village"] || fd.village || "").trim();
      case "Mandal":
      case "mandal":
        return String(cust["Mandal"] || cust["mandal"] || fd["Mandal"] || fd.mandal || "").trim();
      case "Mobile Numb":
      case "Mobile Number":
        return String(cust["Mobile Numb"] || cust["Mobile Number"] || cust.phoneNo || cust.mobileNumber || cust.ownerMob || cust.phNo || fd["Mobile Numb"] || fd["Mobile Number"] || fd.phoneNo || fd.ownerMob || "").trim();
      case "Distict":
      case "District":
        return String(cust["Distict"] || cust["District"] || cust.district || fd["Distict"] || fd["District"] || fd.district || "").trim();
      case "PIN CO":
      case "Pin Code":
        return String(cust["PIN CO"] || cust["PIN CODE"] || cust.pinCode || fd["PIN CO"] || fd["PIN CODE"] || fd.pinCode || "").trim();
      case "DSP Name":
        return String(cust["DSP Name"] || cust.dspName || fd["DSP Name"] || fd.dspName || "").trim();
      case "EXCHAI":
        return String(cust["EXCHAI"] || cust.exchAi || fd["EXCHAI"] || fd.exchAi || "").trim();
      case "EXCHANGE TRACTOR MODELS":
        return String(cust["EXCHANGE TRACTOR MODELS"] || cust.exchangeTractorModels || fd["EXCHANGE TRACTOR MODELS"] || fd.exchangeTractorModels || "").trim();

      // 5-Column Consolidated view keys:
      case "col_cust_id":
        return `${cust["Chassis no"] || cust.chassisNo || fd["Chassis no"] || ""} ${cust["BRANCH"] || cust.branch || fd["BRANCH"] || ""}`;
      case "col_customer_info":
        return `${cust["Customer Name"] || cust.custName || fd["Customer Name"] || ""} ${cust["FATHER NAME"] || cust.fatherName || fd["FATHER NAME"] || ""} ${cust["VILLAGE"] || cust.village || fd["VILLAGE"] || ""} ${cust["Mandal"] || cust.mandal || fd["Mandal"] || ""} ${cust["Mobile Numb"] || cust.phoneNo || fd["Mobile Numb"] || ""}`;
      case "col_tractor_spec":
        return `${cust["Model"] || cust.model || fd["Model"] || ""} ${cust["MODEL TYPE"] || cust.modelType || fd["MODEL TYPE"] || ""} ${cust["Engine No:"] || cust.engineNo || fd["Engine No:"] || ""}`;
      case "col_delivery_date":
        return `${cust["Date of del"] || cust.dateOfDelivery || fd["Date of del"] || ""}`;
      case "col_actions":
        return "";

      default:
        return String(cust[colKey] || fd[colKey] || "").trim();
    }
  };

  const handleToggleFilterValue = (colKey: string, val: string) => {
    setColumnFilters((prev) => {
      const current = prev[colKey] || [];
      const exists = current.includes(val);
      const next = exists ? current.filter((x) => x !== val) : [...current, val];
      if (next.length === 0) {
        const copy = { ...prev };
        delete copy[colKey];
        return copy;
      }
      return { ...prev, [colKey]: next };
    });
    setCurrentPage(1);
  };

  const activeColUniqueValues = useMemo(() => {
    if (!activeFilterCol) return [];
    const valMap = new Map<string, number>();
    customers.forEach((c) => {
      const val = getColDisplayValue(c, activeFilterCol);
      if (val) {
        valMap.set(val, (valMap.get(val) || 0) + 1);
      }
    });
    const arr = Array.from(valMap.entries()).map(([value, count]) => ({ value, count }));
    arr.sort((a, b) => a.value.localeCompare(b.value));
    return arr;
  }, [activeFilterCol, customers]);

  const popupDisplayValues = useMemo(() => {
    if (!filterSearchText.trim()) return activeColUniqueValues;
    const q = filterSearchText.toLowerCase();
    return activeColUniqueValues.filter((item) => item.value.toLowerCase().includes(q));
  }, [activeFilterCol, customers, filterSearchText]);

  const processedCustomers = useMemo(() => {
    let result = [...customers];
    const filterKeys = Object.keys(columnFilters);
    if (filterKeys.length > 0) {
      result = result.filter((cust) => {
        return filterKeys.every((colKey) => {
          const selectedVals = columnFilters[colKey];
          if (!selectedVals || selectedVals.length === 0) return true;
          const val = getColDisplayValue(cust, colKey);
          return selectedVals.includes(val);
        });
      });
    }

    if (sortCol) {
      result.sort((a, b) => {
        const valA = getColDisplayValue(a, sortCol);
        const valB = getColDisplayValue(b, sortCol);
        if (valA === valB) return 0;
        const cmp = valA.localeCompare(valB, undefined, { numeric: true });
        return sortDir === "asc" ? cmp : -cmp;
      });
    }
    return result;
  }, [customers, columnFilters, sortCol, sortDir]);

  const totalPages = Math.ceil(processedCustomers.length / pageSize) || 1;
  const paginatedCustomers = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return processedCustomers.slice(start, start + pageSize);
  }, [processedCustomers, currentPage, pageSize]);

  const handleSaveRow = async (cust: any, idx: number) => {
    const key = getRowKey(cust, idx);
    const chassis = cust["Chassis no"] || cust.chassisNo || cust.chassis || "";
    if (!chassis) {
      alert("Chassis number is required to save customer master data.");
      return;
    }
    setSavingRows((prev) => ({ ...prev, [key]: true }));
    try {
      const draft = rowDrafts[key] || {};
      onSave(chassis, draft);
      setSavedSuccessRows((prev) => ({ ...prev, [key]: true }));
      setTimeout(() => {
        setSavedSuccessRows((prev) => ({ ...prev, [key]: false }));
      }, 2000);
    } catch (err) {
      console.error(err);
    } finally {
      setSavingRows((prev) => ({ ...prev, [key]: false }));
    }
  };

  const handleExportExcel = () => {
    if (processedCustomers.length === 0) {
      alert("No customer records to export.");
      return;
    }
    const ws = XLSX.utils.json_to_sheet(processedCustomers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Master Customer Database");
    XLSX.writeFile(wb, `SriGayathri_Master_Customers_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const filterPopupRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (filterPopupRef.current && !filterPopupRef.current.contains(event.target as Node)) {
        setActiveFilterCol(null);
      }
    };
    if (activeFilterCol) {
      document.addEventListener("mousedown", handleClickOutside);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [activeFilterCol]);

  const cellPadding =
    rowDensity === "compact"
      ? "py-1 px-2 text-[11px]"
      : rowDensity === "spacious"
      ? "py-3 px-3 text-xs"
      : "py-2 px-2.5 text-[11.5px]";

  const detailedCols = [
    { key: "SUPERVISOR", label: "SUPERVISOR" },
    { key: "BRANCH", label: "BRANCH" },
    { key: "SL.NO", label: "SL.NO" },
    { key: "Model", label: "Model" },
    { key: "MODEL TYPE", label: "MODEL TYPE" },
    { key: "Chassis no", label: "Chassis no" },
    { key: "Engine No:", label: "Engine No:" },
    { key: "Date of del", label: "Date of del" },
    { key: "Customer Name", label: "Customer Name" },
    { key: "FATHER NAME", label: "FATHER NAME" },
    { key: "ADDRESS", label: "ADDRESS" },
    { key: "VILLAGE", label: "VILLAGE" },
    { key: "Mandal", label: "Mandal" },
    { key: "Mobile Numb", label: "Mobile Numb" },
    { key: "Distict", label: "Distict" },
    { key: "PIN CO", label: "PIN CO" },
    { key: "DSP Name", label: "DSP Name" },
    { key: "EXCHAI", label: "EXCHAI" },
    { key: "EXCHANGE TRACTOR MODELS", label: "EXCHANGE TRACTOR MODELS" }
  ];

  const renderHeaderCell = (colKey: string, label: string) => {
    const isFiltered = (columnFilters[colKey] || []).length > 0;
    const isSorted = sortCol === colKey;

    return (
      <th
        key={colKey}
        className="bg-purple-900 text-white font-extrabold uppercase tracking-wider text-[10.5px] border-r border-b-2 border-purple-950 p-2.5 select-none whitespace-nowrap group hover:bg-purple-800 transition-colors relative"
      >
        <div className="flex items-center justify-between gap-1.5">
          <button
            type="button"
            onClick={() => {
              if (sortCol === colKey) {
                setSortDir((prev) => (prev === "asc" ? "desc" : "asc"));
              } else {
                setSortCol(colKey);
                setSortDir("asc");
              }
            }}
            className="flex items-center gap-1 hover:text-amber-300 transition-colors text-left font-extrabold flex-1 truncate cursor-pointer"
          >
            <span>{label}</span>
            {isSorted && (
              <span className="text-amber-300">
                {sortDir === "asc" ? <ArrowUp className="w-3 h-3" /> : <ArrowDown className="w-3 h-3" />}
              </span>
            )}
          </button>
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              setActiveFilterCol(activeFilterCol === colKey ? null : colKey);
              setFilterSearchText("");
            }}
            className={`p-1 rounded transition-colors cursor-pointer ${
              isFiltered
                ? "bg-amber-500 text-slate-950 font-black shadow-xs"
                : "text-purple-300 hover:text-white hover:bg-purple-800"
            }`}
            title="Filter column"
          >
            <Filter className="w-3 h-3" />
          </button>
        </div>

        {activeFilterCol === colKey && (
          <div
            ref={filterPopupRef}
            className="absolute left-0 top-full mt-1 w-64 bg-white rounded-xl shadow-2xl border border-slate-300 z-50 p-3 text-xs normal-case font-normal text-slate-800 animate-in fade-in zoom-in-95 duration-100"
          >
            <div className="flex items-center justify-between border-b border-slate-200 pb-2 mb-2 font-bold text-slate-800">
              <span className="flex items-center gap-1.5 text-[11px] text-purple-900 font-extrabold">
                <Filter className="w-3.5 h-3.5 text-purple-600" />
                {label} Filter
              </span>
              <button
                type="button"
                onClick={() => setActiveFilterCol(null)}
                className="text-slate-400 hover:text-slate-600 p-0.5 rounded cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex items-center gap-1 border-b border-slate-150 pb-2 mb-2">
              <button
                type="button"
                onClick={() => {
                  setSortCol(colKey);
                  setSortDir("asc");
                  setActiveFilterCol(null);
                }}
                className="flex-1 py-1 px-1.5 bg-slate-100 hover:bg-purple-50 hover:text-purple-900 rounded font-bold text-[10px] text-center cursor-pointer"
              >
                A → Z Asc
              </button>
              <button
                type="button"
                onClick={() => {
                  setSortCol(colKey);
                  setSortDir("desc");
                  setActiveFilterCol(null);
                }}
                className="flex-1 py-1 px-1.5 bg-slate-100 hover:bg-purple-50 hover:text-purple-900 rounded font-bold text-[10px] text-center cursor-pointer"
              >
                Z → A Desc
              </button>
            </div>
            <div className="relative mb-2">
              <Search className="w-3 h-3 text-slate-400 absolute left-2 top-2" />
              <input
                type="text"
                value={filterSearchText}
                onChange={(e) => setFilterSearchText(e.target.value)}
                placeholder="Search values..."
                className="w-full pl-7 pr-2 py-1 bg-slate-50 border border-slate-200 rounded text-xs outline-none focus:border-purple-600 font-medium"
              />
            </div>
            <div className="max-h-48 overflow-y-auto space-y-0.5 pr-1">
              {popupDisplayValues.map((item, idx) => {
                const isSelected = (columnFilters[colKey] || []).includes(item.value);
                return (
                  <div
                    key={`${item.value}-${idx}`}
                    onClick={() => handleToggleFilterValue(colKey, item.value)}
                    className={`flex items-center justify-between px-2 py-1 rounded cursor-pointer text-xs ${
                      isSelected ? "bg-purple-100 font-bold text-purple-950" : "hover:bg-slate-100 text-slate-700"
                    }`}
                  >
                    <span className="truncate mr-2">{item.value || "(Blank)"}</span>
                    <span className="text-[10px] text-slate-400 font-mono">({item.count})</span>
                  </div>
                );
              })}
            </div>
            {isFiltered && (
              <button
                type="button"
                onClick={() => {
                  setColumnFilters((prev) => {
                    const copy = { ...prev };
                    delete copy[colKey];
                    return copy;
                  });
                  setActiveFilterCol(null);
                }}
                className="w-full mt-2 py-1 bg-red-50 hover:bg-red-100 text-red-700 rounded font-bold text-[10px] text-center cursor-pointer"
              >
                Clear Filter for {label}
              </button>
            )}
          </div>
        )}
      </th>
    );
  };

  return (
    <div className="w-full bg-white border border-slate-200 rounded-2xl shadow-sm p-4 space-y-4">
      {/* Header & Controls bar */}
      <div className="flex flex-wrap items-center justify-between gap-3 bg-purple-50/70 p-3 rounded-xl border border-purple-200">
        <div className="flex items-center gap-2">
          <div className="bg-purple-900 text-white p-2 rounded-xl shadow-xs">
            <TableIcon className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-sm font-black text-purple-950 leading-tight">
              {isTe ? "కస్టమర్ వివరాలు & మాస్టర్ ఎక్సెల్ షీట్" : "Customer Master Excel Spreadsheet View"}
            </h3>
            <p className="text-[11px] text-purple-800 font-medium">
              {isTe
                ? "అన్ని కస్టమర్ రికార్డులు ఎక్సెల్ షీట్ లాగా నేరుగా ఎడిట్ చేసుకోవచ్చు (5-Column & Detailed View)"
                : "Full Excel spreadsheet mode with inline direct cell editing for all master customer records."}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {/* View mode toggle */}
          <div className="flex items-center bg-white p-1 rounded-xl border border-purple-300 shadow-2xs">
            <button
              type="button"
              onClick={() => setViewMode("5col")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                viewMode === "5col"
                  ? "bg-purple-900 text-white shadow-xs"
                  : "text-slate-700 hover:bg-purple-50"
              }`}
            >
              <Columns className="w-3.5 h-3.5" />
              <span>5-Col Master View</span>
            </button>
            <button
              type="button"
              onClick={() => setViewMode("detailed")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                viewMode === "detailed"
                  ? "bg-purple-900 text-white shadow-xs"
                  : "text-slate-700 hover:bg-purple-50"
              }`}
            >
              <TableIcon className="w-3.5 h-3.5" />
              <span>All Columns ({detailedCols.length})</span>
            </button>
          </div>

          {/* Density toggle */}
          <select
            value={rowDensity}
            onChange={(e: any) => setRowDensity(e.target.value)}
            className="text-xs font-bold bg-white text-slate-800 border border-purple-300 rounded-xl px-2.5 py-1.5 outline-none cursor-pointer shadow-2xs"
          >
            <option value="compact">Compact Row</option>
            <option value="normal">Normal Row</option>
            <option value="spacious">Spacious Row</option>
          </select>

          {/* Export Button */}
          <button
            type="button"
            onClick={handleExportExcel}
            className="px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl transition-all shadow-xs flex items-center gap-1.5 cursor-pointer"
          >
            <Download className="w-4 h-4" />
            <span>Export Excel</span>
          </button>

          {Object.keys(columnFilters).length > 0 && (
            <button
              type="button"
              onClick={() => setColumnFilters({})}
              className="px-3 py-1.5 bg-red-50 hover:bg-red-100 text-red-700 font-bold text-xs rounded-xl border border-red-200 transition-all flex items-center gap-1 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Clear Filters ({Object.keys(columnFilters).length})</span>
            </button>
          )}
        </div>
      </div>

      {/* Spreadsheet Table Container */}
      <div className="overflow-x-auto border border-purple-200 rounded-xl shadow-xs bg-white">
        <table className="w-full text-left border-collapse">
          <thead>
            {viewMode === "5col" ? (
              <tr>
                <th className="bg-purple-900 text-white font-extrabold text-[10.5px] p-2.5 border-r border-b-2 border-purple-950 w-12 text-center">
                  #
                </th>
                {renderHeaderCell("Chassis no", isTe ? "1. ఛాసిస్ & బ్రాంచ్ (Chassis & Branch)" : "1. Chassis & Branch")}
                {renderHeaderCell("Customer Name", isTe ? "2. కస్టమర్ & అడ్రస్ (Customer & Address)" : "2. Customer & Address")}
                {renderHeaderCell("Model", isTe ? "3. ట్రాక్టర్ మోడల్ & ఇంజన్ (Tractor Model)" : "3. Tractor & Model")}
                {renderHeaderCell("Date of Delivery", isTe ? "4. డెలివరీ తేదీ (Delivery Date)" : "4. Delivery Date")}
                <th className="bg-purple-900 text-white font-extrabold text-[10.5px] p-2.5 border-b-2 border-purple-950 w-32 text-center">
                  5. Action / Save
                </th>
              </tr>
            ) : (
              <tr>
                <th className="bg-purple-900 text-white font-extrabold text-[10.5px] p-2.5 border-r border-b-2 border-purple-950 w-12 text-center">
                  #
                </th>
                {detailedCols.map((col) => renderHeaderCell(col.key, col.label))}
                <th className="bg-purple-900 text-white font-extrabold text-[10.5px] p-2.5 border-b-2 border-purple-950 w-28 text-center">
                  Save / Action
                </th>
              </tr>
            )}
          </thead>
          <tbody className="divide-y divide-slate-200 font-medium">
            {paginatedCustomers.length === 0 ? (
              <tr>
                <td colSpan={viewMode === "5col" ? 6 : detailedCols.length + 2} className="py-12 text-center text-slate-400">
                  <p className="text-sm font-bold">No customer records found matching current filters.</p>
                </td>
              </tr>
            ) : (
              paginatedCustomers.map((cust, idx) => {
                const globalIdx = (currentPage - 1) * pageSize + idx + 1;
                const rowUniqueIndex = (currentPage - 1) * pageSize + idx;
                const key = getRowKey(cust, rowUniqueIndex);
                const isSaving = savingRows[key];
                const isSaved = savedSuccessRows[key];

                if (viewMode === "5col") {
                  return (
                    <tr key={key} className="hover:bg-purple-50/50 transition-colors group">
                      <td className={`${cellPadding} text-center font-mono text-slate-400 font-bold border-r border-slate-200 bg-slate-50/60`}>
                        {globalIdx}
                      </td>

                      {/* Col 1: Chassis & Branch */}
                      <td className={`${cellPadding} border-r border-slate-200`}>
                        <div className="space-y-1">
                          <input
                            type="text"
                            value={getCustValue(cust, "Chassis no", rowUniqueIndex) || getCustValue(cust, "chassisNo", rowUniqueIndex) || ""}
                            onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Chassis no", e.target.value)}
                            className="w-full font-mono font-bold text-xs bg-slate-50 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-1 outline-none transition-all"
                            placeholder="Chassis No..."
                          />
                          <input
                            type="text"
                            value={getCustValue(cust, "Branch", rowUniqueIndex) || getCustValue(cust, "branch", rowUniqueIndex) || ""}
                            onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Branch", e.target.value)}
                            className="w-full text-[10px] text-purple-900 font-extrabold bg-purple-50/40 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-0.5 outline-none transition-all"
                            placeholder="Branch..."
                          />
                        </div>
                      </td>

                      {/* Col 2: Customer Name & Address */}
                      <td className={`${cellPadding} border-r border-slate-200`}>
                        <div className="space-y-1">
                          <input
                            type="text"
                            value={getCustValue(cust, "Customer Name", rowUniqueIndex) || getCustValue(cust, "custName", rowUniqueIndex) || ""}
                            onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Customer Name", e.target.value)}
                            className="w-full font-black text-slate-950 text-xs bg-slate-50 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-1 outline-none transition-all"
                            placeholder="Customer Name..."
                          />
                          <div className="grid grid-cols-2 gap-1">
                            <input
                              type="text"
                              value={getCustValue(cust, "Village", rowUniqueIndex) || getCustValue(cust, "village", rowUniqueIndex) || ""}
                              onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Village", e.target.value)}
                              className="w-full text-[11px] text-slate-700 font-semibold bg-slate-50 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-0.5 outline-none transition-all"
                              placeholder="Village..."
                            />
                            <input
                              type="text"
                              value={getCustValue(cust, "Mobile Number", rowUniqueIndex) || getCustValue(cust, "phoneNo", rowUniqueIndex) || ""}
                              onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Mobile Number", e.target.value)}
                              className="w-full font-mono text-[11px] text-emerald-700 font-black bg-emerald-50/40 hover:bg-white focus:bg-white border border-transparent hover:border-emerald-300 focus:border-emerald-600 rounded px-1.5 py-0.5 outline-none transition-all"
                              placeholder="Mobile No..."
                            />
                          </div>
                        </div>
                      </td>

                      {/* Col 3: Tractor Model & Engine */}
                      <td className={`${cellPadding} border-r border-slate-200`}>
                        <div className="space-y-1">
                          <input
                            type="text"
                            value={getCustValue(cust, "Model", rowUniqueIndex) || getCustValue(cust, "model", rowUniqueIndex) || ""}
                            onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Model", e.target.value)}
                            className="w-full font-bold text-slate-900 text-xs bg-slate-50 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-1 outline-none transition-all"
                            placeholder="Tractor Model..."
                          />
                          <input
                            type="text"
                            value={getCustValue(cust, "Engine No", rowUniqueIndex) || getCustValue(cust, "engineNo", rowUniqueIndex) || ""}
                            onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Engine No", e.target.value)}
                            className="w-full font-mono text-[10px] text-slate-600 bg-slate-50 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-0.5 outline-none transition-all"
                            placeholder="Engine No..."
                          />
                        </div>
                      </td>

                      {/* Col 4: Delivery Date */}
                      <td className={`${cellPadding} border-r border-slate-200`}>
                        <input
                          type="text"
                          value={getCustValue(cust, "Date of Delivery", rowUniqueIndex) || getCustValue(cust, "dateOfDelivery", rowUniqueIndex) || ""}
                          onChange={(e) => handleFieldChange(cust, rowUniqueIndex, "Date of Delivery", e.target.value)}
                          className="w-full font-mono text-xs text-slate-800 bg-slate-50 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-1.5 outline-none transition-all"
                          placeholder="DD/MM/YYYY..."
                        />
                      </td>

                      {/* Col 5: Actions */}
                      <td className={`${cellPadding} text-center`}>
                        <div className="flex items-center justify-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => handleSaveRow(cust, rowUniqueIndex)}
                            disabled={isSaving}
                            className={`px-2.5 py-1.5 rounded-lg text-xs font-bold transition-all shadow-xs flex items-center gap-1 cursor-pointer ${
                              isSaved
                                ? "bg-emerald-600 text-white"
                                : "bg-purple-900 hover:bg-purple-950 text-white"
                            }`}
                            title="Save Changes"
                          >
                            {isSaved ? <Check className="w-3.5 h-3.5" /> : <Save className="w-3.5 h-3.5" />}
                            <span>{isSaved ? "Saved" : "Save"}</span>
                          </button>
                          {onDelete && (
                            <button
                              type="button"
                              onClick={() => {
                                const ch = cust["Chassis no"] || cust.chassisNo || cust.chassis;
                                if (ch) onDelete(ch);
                              }}
                              className="p-1.5 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                              title="Delete Customer"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                } else {
                  // Detailed mode: all columns editable inline
                  return (
                    <tr key={key} className="hover:bg-purple-50/50 transition-colors group">
                      <td className={`${cellPadding} text-center font-mono text-slate-400 font-bold border-r border-slate-200 bg-slate-50/60`}>
                        {globalIdx}
                      </td>
                      {detailedCols.map((col) => (
                        <td key={col.key} className={`${cellPadding} border-r border-slate-200`}>
                          <input
                            type="text"
                            value={getCustValue(cust, col.key, rowUniqueIndex)}
                            onChange={(e) => handleFieldChange(cust, rowUniqueIndex, col.key, e.target.value)}
                            className="w-full text-xs font-medium bg-slate-50 hover:bg-white focus:bg-white border border-transparent hover:border-purple-300 focus:border-purple-600 rounded px-1.5 py-1 outline-none transition-all"
                          />
                        </td>
                      ))}
                      <td className={`${cellPadding} text-center`}>
                        <div className="flex items-center justify-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => handleSaveRow(cust, rowUniqueIndex)}
                            disabled={isSaving}
                            className={`px-2.5 py-1.5 rounded-lg text-xs font-bold transition-all shadow-xs flex items-center gap-1 cursor-pointer ${
                              isSaved
                                ? "bg-emerald-600 text-white"
                                : "bg-purple-900 hover:bg-purple-950 text-white"
                            }`}
                          >
                            {isSaved ? <Check className="w-3.5 h-3.5" /> : <Save className="w-3.5 h-3.5" />}
                            <span>{isSaved ? "Saved" : "Save"}</span>
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                }
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination Footer */}
      <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-200 text-xs text-slate-600 font-bold">
        <div className="flex items-center gap-2">
          <span>
            Showing {(currentPage - 1) * pageSize + 1} - {Math.min(currentPage * pageSize, processedCustomers.length)} of {processedCustomers.length} customers
          </span>
          <select
            value={pageSize}
            onChange={(e) => {
              setPageSize(Number(e.target.value));
              setCurrentPage(1);
            }}
            className="text-xs border border-purple-300 rounded px-2 py-1 font-bold outline-none cursor-pointer bg-white"
          >
            <option value="15">15 per page</option>
            <option value="25">25 per page</option>
            <option value="50">50 per page</option>
            <option value="100">100 per page</option>
          </select>
        </div>

        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            disabled={currentPage === 1}
            className="px-3 py-1 bg-purple-50 hover:bg-purple-100 disabled:opacity-45 text-purple-950 rounded font-bold cursor-pointer transition-all"
          >
            Previous
          </button>
          <span className="px-3 py-1 bg-purple-900 text-white rounded font-black">
            {currentPage} / {totalPages || 1}
          </span>
          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            disabled={currentPage === totalPages || totalPages === 0}
            className="px-3 py-1 bg-purple-50 hover:bg-purple-100 disabled:opacity-45 text-purple-950 rounded font-bold cursor-pointer transition-all"
          >
            Next
          </button>
        </div>
      </div>
    </div>
  );
};
