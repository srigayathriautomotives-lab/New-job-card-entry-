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
  CheckSquare,
  Square,
  CheckCircle2,
  Phone,
  Maximize2,
  Minimize2,
  Table as TableIcon,
  Columns
} from "lucide-react";
import * as XLSX from "xlsx";

export interface SavedJobCardsExcelTableProps {
  cards: any[];
  allCards: any[];
  language?: "te" | "en";
  onView: (card: any) => void;
  onEdit: (card: any) => void;
  onPrint: (card: any) => void;
  onSave: (cardId: string, updatedFields: any) => Promise<void>;
  onDelete?: (cardId: string) => void;
  selectedIds: string[];
  onToggleSelect: (cardId: string) => void;
  onSelectAll: (ids: string[]) => void;
  mechanicsList?: string[];
  branchesList?: string[];
  serviceTypes?: string[];
}

export const SavedJobCardsExcelTable: React.FC<SavedJobCardsExcelTableProps> = ({
  cards,
  allCards,
  language = "te",
  onView,
  onEdit,
  onPrint,
  onSave,
  selectedIds,
  onToggleSelect,
  onSelectAll,
  mechanicsList = ["Srinu", "Rambabu", "Siva", "Prasad", "Nani", "Kiran"],
  branchesList = ["Vuyyuru", "Vijayawada", "Machilipatnam", "Gudivada", "Nuzvid", "Tiruvuru"],
  serviceTypes = [
    "1st Free",
    "2nd Free",
    "3rd Free",
    "4th Free",
    "5th Free",
    "6th Free",
    "7th Free",
    "8th Free",
    "9th Free",
    "10th Free",
    "Paid Service",
    "Paid Repairs",
    "Running Repair",
    "Under Wty Repairs",
    "Major Overhaul",
    "PDI",
    "Breakdown Service",
    "Campaign"
  ]
}) => {
  const isTe = language === "te";

  // Table view mode: "5col" (5 Master Columns Excel) or "detailed" (All 26 Excel Sheet Columns)
  const [viewMode, setViewMode] = useState<"5col" | "detailed">("5col");
  const [rowDensity, setRowDensity] = useState<"compact" | "normal" | "spacious">("normal");

  // Pagination
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(25);

  // Sorting
  const [sortCol, setSortCol] = useState<string>("complaintDate");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("desc");

  // Column Filters State: Record<columnKey, Set<string>>
  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [filterSearchText, setFilterSearchText] = useState<string>("");

  // Saving state indicator per row: cardId -> boolean
  const [savingRows, setSavingRows] = useState<Record<string, boolean>>({});
  const [savedSuccessRows, setSavedSuccessRows] = useState<Record<string, boolean>>({});

  // State to store draft changes for the rows in direct spreadsheet editing
  const [rowDrafts, setRowDrafts] = useState<Record<string, any>>({});

  const toInputDate = (val: any): string => {
    if (!val) return "";
    const s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
      const [d, m, y] = s.split("/");
      return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
    }
    if (/^\d{2}-\d{2}-\d{4}$/.test(s)) {
      const [d, m, y] = s.split("-");
      return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
    }
    const dateObj = new Date(val);
    if (!isNaN(dateObj.getTime())) {
      return dateObj.toISOString().slice(0, 10);
    }
    return "";
  };

  const getRowValue = (card: any, fieldKey: string) => {
    if (rowDrafts[card.id] && rowDrafts[card.id][fieldKey] !== undefined) {
      return rowDrafts[card.id][fieldKey];
    }
    return card[fieldKey] ?? "";
  };

  const handleRowFieldChange = (cardId: string, fieldKey: string, value: any) => {
    setRowDrafts((prev) => ({
      ...prev,
      [cardId]: {
        ...(prev[cardId] || {}),
        [fieldKey]: value,
      },
    }));
  };

  // Formatting date helper
  const formatDate = (val: any) => {
    if (!val) return "—";
    const s = String(val).trim();
    if (!s) return "—";
    if (s.includes("T")) return s.split("T")[0];
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const [y, m, d] = s.split("-");
      return `${d}/${m}/${y}`;
    }
    return s;
  };

  // Helper to extract value from a card given a column key
  const getCardColValue = (card: any, colKey: string): string => {
    switch (colKey) {
      case "status":
        return card.status === "Closed" || (card.actualClosedDate && card.status !== "Open") ? "Closed" : "Open";
      case "jobNo":
        return String(card.jobNo || card.onlineJobCardNo || "").trim();
      case "onlineJobCardNo":
        return String(card.onlineJobCardNo || card.jobNo || "").trim();
      case "complaintDate":
        return formatDate(card.complaintDate);
      case "jobDate":
        return formatDate(card.jobDate || card.dateTimeIn || card.complaintDate);
      case "branch":
        return String(card.branch || "").trim();
      case "historyFileNo":
        return String(card.historyFileNo || card.fileNo || "").trim();
      case "model":
        return String(card.model || "").trim();
      case "modelType":
        return String(card.modelType || "").trim();
      case "chassisNo":
        return String(card.chassisNo || "").trim();
      case "engineNo":
        return String(card.engineNo || "").trim();
      case "dateOfDelivery":
        return formatDate(card.dateOfDelivery || card.installDate);
      case "custName":
        return String(card.custName || "").trim();
      case "fatherName":
        return String(card.fatherName || "").trim();
      case "village":
        return String(card.village || "").trim();
      case "mandal":
        return String(card.mandal || "").trim();
      case "phoneNo":
        return String(card.ownerMob || card.phNo || "").trim();
      case "hourMeter":
        return String(card.hourMeter || card.hrsRun || 0);
      case "serviceType":
        return String(card.serviceType || "").trim();
      case "freeServiceList":
        return String(card.freeServiceList || "").trim();
      case "extraRepairs":
        return String(card.extraRepairs || "").trim();
      case "actualClosedDate":
        return formatDate(card.actualClosedDate || card.dateTimeOut);
      case "mechanic":
        return String(card.mechanic || card.technicianName || "").trim();
      case "serviceLocation":
        return String(card.serviceLocation || card.servicePlace || "").trim();
      case "billNo":
        return String(card.billNo || "").trim();
      case "reasonsForAnalysis":
        return String(card.reasonsForAnalysis || card.problemDescription || "").trim();
      case "telecalling":
        return String(card.telecalling || "").trim();

      // Consolidated 5-column keys:
      case "col_job_dates":
        return `${card.jobNo || ""} ${card.onlineJobCardNo || ""} ${card.branch || ""} ${formatDate(card.complaintDate)}`;
      case "col_customer":
        return `${card.custName || ""} ${card.fatherName || ""} ${card.village || ""} ${card.mandal || ""} ${card.ownerMob || card.phNo || ""}`;
      case "col_tractor":
        return `${card.model || ""} ${card.modelType || ""} ${card.chassisNo || ""} ${card.engineNo || ""}`;
      case "col_service":
        return `${card.hourMeter || ""} ${card.serviceType || ""} ${card.freeServiceList || ""} ${card.extraRepairs || ""}`;
      case "col_closure":
        return `${card.actualClosedDate || ""} ${card.mechanic || ""} ${card.billNo || ""} ${card.serviceLocation || ""}`;

      default:
        return String(card[colKey] || "").trim();
    }
  };

  // Filter dropdown handler logic
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

  const handleClearColumnFilter = (colKey: string) => {
    setColumnFilters((prev) => {
      const copy = { ...prev };
      delete copy[colKey];
      return copy;
    });
    setCurrentPage(1);
  };

  const handleClearAllFilters = () => {
    setColumnFilters({});
    setFilterSearchText("");
    setActiveFilterCol(null);
    setCurrentPage(1);
  };

  // Compute unique values for active column filter
  const activeColUniqueValues = useMemo(() => {
    if (!activeFilterCol) return [];
    const valMap = new Map<string, number>();
    allCards.forEach((c) => {
      const val = getCardColValue(c, activeFilterCol);
      if (val) {
        valMap.set(val, (valMap.get(val) || 0) + 1);
      }
    });
    const arr = Array.from(valMap.entries()).map(([value, count]) => ({
      value,
      count
    }));
    arr.sort((a, b) => a.value.localeCompare(b.value));
    return arr;
  }, [activeFilterCol, allCards]);

  // Filtered values inside filter popup by filterSearchText
  const popupDisplayValues = useMemo(() => {
    if (!filterSearchText.trim()) return activeColUniqueValues;
    const q = filterSearchText.toLowerCase();
    return activeColUniqueValues.filter((item) =>
      item.value.toLowerCase().includes(q)
    );
  }, [activeColUniqueValues, filterSearchText]);

  // Apply column filters and sorting to raw cards
  const processedCards = useMemo(() => {
    let result = [...cards];

    // Apply active column filters
    const filterKeys = Object.keys(columnFilters);
    if (filterKeys.length > 0) {
      result = result.filter((card) => {
        return filterKeys.every((colKey) => {
          const selectedVals = columnFilters[colKey];
          if (!selectedVals || selectedVals.length === 0) return true;
          const cardVal = getCardColValue(card, colKey);
          return selectedVals.includes(cardVal);
        });
      });
    }

    // Apply sorting
    if (sortCol) {
      result.sort((a, b) => {
        const valA = getCardColValue(a, sortCol);
        const valB = getCardColValue(b, sortCol);
        if (valA === valB) return 0;
        const cmp = valA.localeCompare(valB, undefined, { numeric: true });
        return sortDir === "asc" ? cmp : -cmp;
      });
    }

    return result;
  }, [cards, columnFilters, sortCol, sortDir]);

  // Pagination slicing
  const totalPages = Math.ceil(processedCards.length / pageSize) || 1;
  const paginatedCards = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return processedCards.slice(start, start + pageSize);
  }, [processedCards, currentPage, pageSize]);

  // Toggle status inline and trigger save
  const handleToggleStatus = async (card: any, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    const currentStatus = card.status === "Closed" || (card.actualClosedDate && card.status !== "Open") ? "Closed" : "Open";
    const newStatus = currentStatus === "Open" ? "Closed" : "Open";
    const updated: any = { status: newStatus };
    if (newStatus === "Closed" && !card.actualClosedDate) {
      updated.actualClosedDate = new Date().toISOString().split("T")[0];
    }

    setSavingRows((prev) => ({ ...prev, [card.id]: true }));
    try {
      await onSave(card.id, updated);
      setSavedSuccessRows((prev) => ({ ...prev, [card.id]: true }));
      setTimeout(() => {
        setSavedSuccessRows((prev) => ({ ...prev, [card.id]: false }));
      }, 2000);
    } catch (err) {
      console.error(err);
    } finally {
      setSavingRows((prev) => ({ ...prev, [card.id]: false }));
    }
  };

  // Save row button click
  const handleSaveRow = async (card: any) => {
    setSavingRows((prev) => ({ ...prev, [card.id]: true }));
    try {
      const draft = rowDrafts[card.id] || {};
      const payload: any = {
        ...draft
      };
      // Auto-sync field mappings
      if (draft.jobNo !== undefined) {
        payload.onlineJobCardNo = draft.jobNo;
      }
      if (draft.hourMeter !== undefined) {
        payload.hrsRun = draft.hourMeter;
      }
      if (draft.extraRepairs !== undefined) {
        payload.problemDescription = draft.extraRepairs;
      }
      if (draft.actualClosedDate !== undefined) {
        payload.dateTimeOut = draft.actualClosedDate;
        payload.status = "Closed"; // Auto set to Closed when closed date is provided
      }
      if (draft.mechanic !== undefined) {
        payload.technicianName = draft.mechanic;
      }
      if (draft.phoneNo !== undefined) {
        payload.ownerMob = draft.phoneNo;
        payload.phNo = draft.phoneNo;
      }

      await onSave(card.id, payload);
      setSavedSuccessRows((prev) => ({ ...prev, [card.id]: true }));
      setTimeout(() => {
        setSavedSuccessRows((prev) => ({ ...prev, [card.id]: false }));
      }, 2000);
    } catch (err) {
      console.error(err);
    } finally {
      setSavingRows((prev) => ({ ...prev, [card.id]: false }));
    }
  };

  // Quick export filtered table to Excel
  const handleExportFilteredExcel = () => {
    if (processedCards.length === 0) {
      alert("No records to export.");
      return;
    }
    const exportData = processedCards.map((c, idx) => ({
      "Sl No": idx + 1,
      Status: c.status === "Closed" ? "Closed" : "Open",
      "Job card": c.jobNo || c.onlineJobCardNo || "",
      "COMPLAINT DATE": formatDate(c.complaintDate),
      "ONLINE JOB CARD NO": c.onlineJobCardNo || c.jobNo || "",
      "JOB CARD OPEN DATE": formatDate(c.jobDate || c.jobOpenDate || c.dateTimeIn),
      BRANCH: c.branch || "",
      "HISTORY FILE NO.": c.historyFileNo || c.fileNo || "",
      "Tractor model": c.model || "",
      "MODEL TYPE": c.modelType || "",
      "CHASSIS NO": c.chassisNo || "",
      "Eng Sr no": c.engineNo || "",
      "Date of Delivery": formatDate(c.dateOfDelivery || c.installDate),
      "Customer name": c.custName || "",
      FATHER: c.fatherName || "",
      ADDRESS: c.custAddr || c.address || "",
      Village: c.village || "",
      Mandal: c.mandal || "",
      "Phone No": c.ownerMob || c.phNo || "",
      "Hrs Run": c.hourMeter || c.hrsRun || 0,
      "Type of Service": c.serviceType || "",
      "FREE SERVICE LIST": c.freeServiceList || "",
      "EXTRA REPAIRS DONE": c.extraRepairs || "",
      "ACTUAL CLOSED DATE": formatDate(c.actualClosedDate || c.dateTimeOut),
      "TECHNICIAN NAME": c.mechanic || c.technicianName || "",
      "Service place": c.serviceLocation || c.servicePlace || "",
      "BILL NO.": c.billNo || "",
      "REASONS FOR ANALYSIS": c.reasonsForAnalysis || c.problemDescription || "",
      TELECALLING: c.telecalling || ""
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Saved Job Cards");
    XLSX.writeFile(wb, `SriGayathri_Saved_JobCards_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  // Close filter dropdown on outside click
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

  // Row padding based on density
  const cellPadding =
    rowDensity === "compact"
      ? "py-1 px-2 text-[11px]"
      : rowDensity === "spacious"
      ? "py-3 px-3 text-xs"
      : "py-2 px-2.5 text-[11.5px]";

  // Detailed 26 Columns definition
  const detailedColumns = [
    { key: "jobNo", label: isTe ? "జాబ్ కార్డ్ (Job Card)" : "Job Card No", width: "w-28" },
    { key: "complaintDate", label: isTe ? "కంప్లైంట్ తేదీ" : "Complaint Date", width: "w-28" },
    { key: "onlineJobCardNo", label: isTe ? "ఆన్‌లైన్ జేసీ నెం." : "Online JC No", width: "w-32" },
    { key: "jobDate", label: isTe ? "ఓపెన్ తేదీ" : "Open Date", width: "w-28" },
    { key: "branch", label: isTe ? "బ్రాంచ్ (Branch)" : "Branch", width: "w-28" },
    { key: "historyFileNo", label: isTe ? "ఫైల్ నెం." : "File No", width: "w-24" },
    { key: "model", label: isTe ? "మోడల్ (Model)" : "Model", width: "w-32" },
    { key: "modelType", label: isTe ? "మోడల్ రకం" : "Model Type", width: "w-24" },
    { key: "chassisNo", label: isTe ? "ఛాసిస్ నెం." : "Chassis No", width: "w-40" },
    { key: "engineNo", label: isTe ? "ఇంజన్ నెం." : "Engine No", width: "w-32" },
    { key: "dateOfDelivery", label: isTe ? "డెలివరీ తేదీ" : "Delivery Date", width: "w-28" },
    { key: "custName", label: isTe ? "కస్టమర్ పేరు" : "Customer Name", width: "w-40" },
    { key: "fatherName", label: isTe ? "తండ్రి పేరు" : "Father's Name", width: "w-36" },
    { key: "village", label: isTe ? "గ్రామం" : "Village", width: "w-32" },
    { key: "mandal", label: isTe ? "మండలం" : "Mandal", width: "w-28" },
    { key: "phoneNo", label: isTe ? "మొబైల్ ఫోన్" : "Phone No", width: "w-32" },
    { key: "hourMeter", label: isTe ? "గంటలు (Hrs)" : "Hrs Run", width: "w-20" },
    { key: "serviceType", label: isTe ? "సర్వీస్ రకం" : "Service Type", width: "w-28" },
    { key: "freeServiceList", label: isTe ? "ఉచిత సర్వీస్" : "Free Service", width: "w-28" },
    { key: "extraRepairs", label: isTe ? "అదనపు రిపేర్లు" : "Extra Repairs", width: "w-48" },
    { key: "actualClosedDate", label: isTe ? "క్లోజ్డ్ తేదీ" : "Closed Date", width: "w-28" },
    { key: "mechanic", label: isTe ? "టెక్నీషియన్" : "Technician", width: "w-36" },
    { key: "serviceLocation", label: isTe ? "సర్వీస్ ప్రదేశం" : "Location", width: "w-28" },
    { key: "billNo", label: isTe ? "బిల్ నెం." : "Bill No", width: "w-24" },
    { key: "reasonsForAnalysis", label: isTe ? "కారణాలు / విశ్లేషణ" : "Analysis / Problem", width: "w-52" },
    { key: "telecalling", label: isTe ? "టెలికాలింగ్" : "Telecalling", width: "w-36" }
  ];

  // Helper Header Cell component with Excel-style AutoFilter Icon
  const renderHeaderFilterCell = (colKey: string, label: string, widthClass?: string) => {
    const isFiltered = (columnFilters[colKey] || []).length > 0;
    const isSorted = sortCol === colKey;

    return (
      <th
        key={colKey}
        className={`bg-slate-100 text-slate-800 font-extrabold uppercase tracking-wider text-[10.5px] border-r border-b-2 border-slate-300 p-2 select-none whitespace-nowrap group hover:bg-slate-200 transition-colors relative ${
          widthClass || ""
        }`}
      >
        <div className="flex items-center justify-between gap-1.5">
          {/* Header Title with Click-to-Sort */}
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
            className="flex items-center gap-1 hover:text-emerald-800 transition-colors text-left font-extrabold flex-1 truncate cursor-pointer"
            title="Click to sort"
          >
            <span>{label}</span>
            {isSorted && (
              <span className="text-emerald-700">
                {sortDir === "asc" ? <ArrowUp className="w-3 h-3" /> : <ArrowDown className="w-3 h-3" />}
              </span>
            )}
          </button>

          {/* Excel AutoFilter Dropdown Button */}
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              if (activeFilterCol === colKey) {
                setActiveFilterCol(null);
              } else {
                setActiveFilterCol(colKey);
                setFilterSearchText("");
              }
            }}
            className={`p-1 rounded transition-colors cursor-pointer ${
              isFiltered
                ? "bg-emerald-600 text-white shadow-xs hover:bg-emerald-700"
                : "text-slate-400 hover:text-slate-700 hover:bg-slate-300/60"
            }`}
            title={`Filter by ${label}`}
          >
            <Filter className="w-3 h-3" />
          </button>
        </div>

        {/* Excel Filter Popup */}
        {activeFilterCol === colKey && (
          <div
            ref={filterPopupRef}
            className="absolute left-0 top-full mt-1 w-64 bg-white rounded-xl shadow-2xl border border-slate-300 z-50 p-3 text-xs normal-case font-normal text-slate-800 animate-in fade-in zoom-in-95 duration-100"
          >
            <div className="flex items-center justify-between border-b border-slate-200 pb-2 mb-2 font-bold text-slate-800">
              <span className="flex items-center gap-1.5 text-[11px] text-emerald-800 font-extrabold">
                <Filter className="w-3.5 h-3.5 text-emerald-600" />
                {label} {isTe ? "ఫిల్టర్" : "Filter"}
              </span>
              <button
                type="button"
                onClick={() => setActiveFilterCol(null)}
                className="text-slate-400 hover:text-slate-600 p-0.5 rounded cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Quick Sort Options */}
            <div className="flex items-center gap-1 border-b border-slate-150 pb-2 mb-2">
              <button
                type="button"
                onClick={() => {
                  setSortCol(colKey);
                  setSortDir("asc");
                  setActiveFilterCol(null);
                }}
                className="flex-1 py-1 px-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-800 rounded font-bold text-[10px] text-center transition-colors cursor-pointer"
              >
                A → Z {isTe ? "ఆరోహణ" : "Asc"}
              </button>
              <button
                type="button"
                onClick={() => {
                  setSortCol(colKey);
                  setSortDir("desc");
                  setActiveFilterCol(null);
                }}
                className="flex-1 py-1 px-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-800 rounded font-bold text-[10px] text-center transition-colors cursor-pointer"
              >
                Z → A {isTe ? "అవరోహణ" : "Desc"}
              </button>
            </div>

            {/* Live Search inside filter values */}
            <div className="relative mb-2">
              <Search className="w-3 h-3 text-slate-400 absolute left-2 top-2" />
              <input
                type="text"
                value={filterSearchText}
                onChange={(e) => setFilterSearchText(e.target.value)}
                placeholder={isTe ? "వెతకండి..." : "Search values..."}
                className="w-full bg-slate-50 border border-slate-200 rounded pl-7 pr-2 py-1 text-[11px] outline-none focus:border-emerald-500 focus:bg-white"
                autoFocus
              />
            </div>

            {/* Checkbox List */}
            <div className="max-h-40 overflow-y-auto space-y-1 pr-1 mb-2.5 divide-y divide-slate-100">
              {popupDisplayValues.length === 0 ? (
                <div className="text-center py-4 text-slate-400 text-[11px]">
                  {isTe ? "విలువలు లేవు" : "No matching values"}
                </div>
              ) : (
                popupDisplayValues.map((item) => {
                  const isChecked = (columnFilters[colKey] || []).includes(item.value);
                  return (
                    <label
                      key={item.value}
                      className="flex items-center justify-between p-1 hover:bg-emerald-50/50 rounded cursor-pointer text-[11px]"
                    >
                      <div className="flex items-center gap-2 truncate pr-1">
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => handleToggleFilterValue(colKey, item.value)}
                          className="w-3.5 h-3.5 rounded text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                        />
                        <span className="truncate text-slate-800 font-medium">{item.value}</span>
                      </div>
                      <span className="text-[10px] font-mono text-slate-400 bg-slate-100 px-1 py-0.5 rounded">
                        {item.count}
                      </span>
                    </label>
                  );
                })
              )}
            </div>

            {/* Actions: Clear / Done */}
            <div className="flex items-center justify-between pt-2 border-t border-slate-200 text-[10px]">
              <button
                type="button"
                onClick={() => handleClearColumnFilter(colKey)}
                className="text-rose-600 hover:text-rose-800 font-bold hover:underline cursor-pointer"
              >
                {isTe ? "ఫిల్టర్ తొలగించు" : "Clear Filter"}
              </button>
              <button
                type="button"
                onClick={() => setActiveFilterCol(null)}
                className="px-3 py-1 bg-emerald-700 hover:bg-emerald-800 text-white font-bold rounded shadow-xs cursor-pointer"
              >
                {isTe ? "సరే (OK)" : "OK"}
              </button>
            </div>
          </div>
        )}
      </th>
    );
  };

  return (
    <div className="space-y-2.5 flex flex-col min-h-0 bg-white rounded-xl border border-slate-300 shadow-sm p-3">
      {/* EXCEL SPREADSHEET TOOLBAR */}
      <div className="bg-emerald-800 text-white p-2.5 rounded-lg flex flex-wrap items-center justify-between gap-2.5 shadow-sm">
        <div className="flex items-center gap-2.5 flex-wrap">
          {/* Excel Brand & Title */}
          <div className="flex items-center gap-2 pr-2 border-r border-emerald-700">
            <div className="w-7 h-7 bg-white text-emerald-800 rounded font-black flex items-center justify-center text-xs shadow-inner">
              📊
            </div>
            <div>
              <div className="text-xs font-black tracking-wide flex items-center gap-1.5">
                <span>{isTe ? "సేవ్ చేసిన జాబ్ కార్డ్స్ ఎక్సెల్ షీట్" : "Saved Job Cards Spreadsheet"}</span>
              </div>
              <div className="text-[10px] text-emerald-200 font-mono">
                {processedCards.length} {isTe ? "కార్డులు ఫిల్టర్ అయ్యాయి" : "cards listed"}
              </div>
            </div>
          </div>

          {/* View Mode Toggle: 5 Master Columns vs Detailed 26 Columns */}
          <div className="inline-flex rounded-md bg-emerald-900/60 p-0.5 border border-emerald-700">
            <button
              type="button"
              onClick={() => setViewMode("5col")}
              className={`px-2.5 py-1 rounded text-xs font-extrabold transition-all cursor-pointer flex items-center gap-1.5 ${
                viewMode === "5col"
                  ? "bg-white text-emerald-900 shadow-xs"
                  : "text-emerald-100 hover:text-white"
              }`}
              title="View as 5 Master Columns instead of 5 boxes"
            >
              <Columns className="w-3.5 h-3.5" />
              <span>{isTe ? "5-కాలమ్స్ ఎక్సెల్ వ్యూ" : "5 Master Columns"}</span>
            </button>

            <button
              type="button"
              onClick={() => setViewMode("detailed")}
              className={`px-2.5 py-1 rounded text-xs font-extrabold transition-all cursor-pointer flex items-center gap-1.5 ${
                viewMode === "detailed"
                  ? "bg-white text-emerald-900 shadow-xs"
                  : "text-emerald-100 hover:text-white"
              }`}
              title="View all 26 detailed Excel columns"
            >
              <TableIcon className="w-3.5 h-3.5" />
              <span>{isTe ? "పూర్తి ఎక్సెల్ షీట్ (26 Cols)" : "Full 26-Cols Sheet"}</span>
            </button>
          </div>

          {/* Row Height / Density */}
          <div className="flex items-center gap-1 border-l border-emerald-700 pl-2">
            <span className="text-[10px] font-bold text-emerald-200">
              {isTe ? "లైన్ సైజు:" : "Row Height:"}
            </span>
            <div className="inline-flex rounded bg-emerald-900/60 p-0.5 border border-emerald-700 text-[10px]">
              <button
                type="button"
                onClick={() => setRowDensity("compact")}
                className={`px-1.5 py-0.5 rounded font-bold cursor-pointer ${
                  rowDensity === "compact" ? "bg-emerald-600 text-white" : "text-emerald-200 hover:text-white"
                }`}
              >
                🤏 Compact
              </button>
              <button
                type="button"
                onClick={() => setRowDensity("normal")}
                className={`px-1.5 py-0.5 rounded font-bold cursor-pointer ${
                  rowDensity === "normal" ? "bg-emerald-600 text-white" : "text-emerald-200 hover:text-white"
                }`}
              >
                ↔️ Normal
              </button>
              <button
                type="button"
                onClick={() => setRowDensity("spacious")}
                className={`px-1.5 py-0.5 rounded font-bold cursor-pointer ${
                  rowDensity === "spacious" ? "bg-emerald-600 text-white" : "text-emerald-200 hover:text-white"
                }`}
              >
                ↕️ Spacious
              </button>
            </div>
          </div>
        </div>

        {/* Right Controls: Clear Filters, Export XLSX */}
        <div className="flex items-center gap-2">
          {Object.keys(columnFilters).length > 0 && (
            <button
              type="button"
              onClick={handleClearAllFilters}
              className="flex items-center gap-1 px-2 py-1 bg-amber-400 hover:bg-amber-300 text-amber-950 rounded text-xs font-bold transition-colors cursor-pointer shadow-xs"
              title="Reset all active column filters"
            >
              <RotateCcw className="w-3 h-3" />
              <span>{isTe ? "ఫిల్టర్లు రీసెట్" : "Reset Filters"} ({Object.keys(columnFilters).length})</span>
            </button>
          )}

          <button
            type="button"
            onClick={handleExportFilteredExcel}
            className="flex items-center gap-1.5 px-3 py-1 bg-white hover:bg-emerald-50 text-emerald-900 rounded-lg text-xs font-extrabold transition-colors shadow-xs cursor-pointer"
            title="Export filtered records to Excel"
          >
            <Download className="w-3.5 h-3.5 text-emerald-700" />
            <span>{isTe ? "ఎక్సెల్ డౌన్‌లోడ్" : "Export Excel"}</span>
          </button>
        </div>
      </div>

      {/* ACTIVE COLUMN FILTERS BADGES STRIP */}
      {Object.keys(columnFilters).length > 0 && (
        <div className="flex items-center gap-1.5 flex-wrap p-2 bg-emerald-50 rounded-lg border border-emerald-200 text-xs">
          <span className="font-extrabold text-emerald-900 flex items-center gap-1 text-[11px]">
            <Filter className="w-3.5 h-3.5 text-emerald-700" />
            {isTe ? "యాక్టివ్ ఫిల్టర్లు:" : "Active Filters:"}
          </span>
          {Object.entries(columnFilters).map(([colKey, vals]) => (
            <span
              key={colKey}
              className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-white border border-emerald-300 text-emerald-900 font-bold text-[11px] shadow-2xs"
            >
              <span className="text-emerald-700 font-semibold">{colKey}:</span>
              <span>{vals.join(", ")}</span>
              <button
                type="button"
                onClick={() => handleClearColumnFilter(colKey)}
                className="text-slate-400 hover:text-rose-600 p-0.5 cursor-pointer"
              >
                <X className="w-3 h-3" />
              </button>
            </span>
          ))}
          <button
            type="button"
            onClick={handleClearAllFilters}
            className="text-[10.5px] font-bold text-rose-600 hover:text-rose-800 underline ml-1 cursor-pointer"
          >
            {isTe ? "అన్నీ తీసివేయి" : "Clear All"}
          </button>
        </div>
      )}

      {/* SPREADSHEET TABLE CONTAINER */}
      <div className="overflow-auto border border-slate-300 rounded-lg shadow-xs bg-white max-h-[calc(100vh-320px)] relative">
        <table className="w-full text-left text-slate-700 border-collapse border border-slate-300 font-sans">
          {/* HEADER ROW */}
          <thead className="sticky top-0 z-30 shadow-xs">
            <tr>
              {/* Col 1: Sl No & Selection */}
              <th className="bg-slate-100 text-slate-800 font-extrabold uppercase tracking-wider text-[10.5px] border-r border-b-2 border-slate-300 p-2 text-center select-none w-14">
                <div className="flex items-center justify-center gap-1">
                  <input
                    type="checkbox"
                    checked={selectedIds.length === paginatedCards.length && paginatedCards.length > 0}
                    onChange={(e) => {
                      if (e.target.checked) {
                        onSelectAll(paginatedCards.map((c) => c.id));
                      } else {
                        onSelectAll([]);
                      }
                    }}
                    className="w-3.5 h-3.5 rounded text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                  />
                  <span>#</span>
                </div>
              </th>

              {/* Col 2: STATUS (LINE KI MODATLO CLOSED / OPEN STATUS) */}
              {renderHeaderFilterCell("status", isTe ? "స్టేటస్ (Status)" : "Status", "w-24 text-center")}

              {/* MIDDLE COLUMNS: EITHER 5 MASTER COLUMNS OR 26 DETAILED COLS */}
              {viewMode === "5col" ? (
                <>
                  {/* Master Col 1: Job Details & Dates */}
                  {renderHeaderFilterCell("branch", isTe ? "1. జాబ్ & తేదీల వివరాలు (Job & Dates)" : "1. Job & Dates", "w-64")}

                  {/* Master Col 2: Customer Details */}
                  {renderHeaderFilterCell("village", isTe ? "2. కస్టమర్ వివరాలు (Customer Details)" : "2. Customer Details", "w-64")}

                  {/* Master Col 3: Tractor Details */}
                  {renderHeaderFilterCell("model", isTe ? "3. ట్రాక్టర్ వివరాలు (Tractor Details)" : "3. Tractor Details", "w-60")}

                  {/* Master Col 4: Service & Repairs */}
                  {renderHeaderFilterCell("serviceType", isTe ? "4. సర్వీస్ & రిపేర్లు (Service & Repairs)" : "4. Service & Repairs", "w-64")}

                  {/* Master Col 5: Closure & Technician */}
                  {renderHeaderFilterCell("mechanic", isTe ? "5. క్లోజర్ & టెక్నీషియన్ (Closure & Mechanic)" : "5. Closure & Mechanic", "w-64")}
                </>
              ) : (
                detailedColumns.map((col) =>
                  renderHeaderFilterCell(col.key, col.label, col.width)
                )
              )}

              {/* LAST COLUMN: ACTIONS (EYE, EDIT, PRINT, SAVE OPTIONS) */}
              <th className="bg-slate-100 text-slate-800 font-extrabold uppercase tracking-wider text-[10.5px] border-b-2 border-slate-300 p-2 text-center sticky right-0 z-40 bg-slate-100 shadow-[-4px_0_6px_-2px_rgba(0,0,0,0.06)] w-36">
                {isTe ? "చర్యలు (Actions)" : "Actions"}
              </th>
            </tr>
          </thead>

          {/* TABLE BODY */}
          <tbody className="divide-y divide-slate-200">
            {paginatedCards.length === 0 ? (
              <tr>
                <td
                  colSpan={viewMode === "5col" ? 8 : detailedColumns.length + 3}
                  className="p-8 text-center text-slate-400 font-bold bg-slate-50"
                >
                  <div className="flex flex-col items-center gap-2">
                    <span className="text-2xl">🔍</span>
                    <span>
                      {isTe
                        ? "ఫిల్టర్‌కు సరిపోలే జాబ్ కార్డ్స్ లేవు."
                        : "No job cards match your active filter criteria."}
                    </span>
                    {Object.keys(columnFilters).length > 0 && (
                      <button
                        type="button"
                        onClick={handleClearAllFilters}
                        className="px-3 py-1 bg-emerald-700 text-white rounded-lg text-xs font-bold mt-1 cursor-pointer hover:bg-emerald-800"
                      >
                        {isTe ? "ఫిల్టర్లు క్లియర్ చేయండి" : "Clear All Filters"}
                      </button>
                    )}
                  </div>
                </td>
              </tr>
            ) : (
              paginatedCards.map((card, idx) => {
                const rowNum = (currentPage - 1) * pageSize + idx + 1;
                const isSelected = selectedIds.includes(card.id);
                const isClosed = card.status === "Closed" || (card.actualClosedDate && card.status !== "Open");
                const isSaving = savingRows[card.id];
                const isSavedSuccess = savedSuccessRows[card.id];

                return (
                  <tr
                    key={card.id || idx}
                    className={`transition-colors border-b border-slate-200 hover:bg-emerald-50/40 ${
                      isSelected
                        ? "bg-emerald-100/60"
                        : idx % 2 === 0
                        ? "bg-white"
                        : "bg-slate-50/50"
                    }`}
                  >
                    {/* Col 1: Sl No & Checkbox */}
                    <td className={`${cellPadding} text-center font-bold border-r border-slate-200 select-none`}>
                      <div className="flex items-center justify-center gap-1.5">
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => onToggleSelect(card.id)}
                          className="w-3.5 h-3.5 rounded text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                        />
                        <span className="font-mono text-slate-500 text-[11px]">{rowNum}</span>
                      </div>
                    </td>

                    {/* Col 2: STATUS (CLOSED / OPEN STATUS AT START OF LINE) */}
                    <td className={`${cellPadding} text-center border-r border-slate-200`}>
                      <button
                        type="button"
                        onClick={(e) => handleToggleStatus(card, e)}
                        disabled={isSaving}
                        className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider border transition-all cursor-pointer shadow-2xs hover:scale-105 active:scale-95 ${
                          isClosed
                            ? "bg-slate-100 text-slate-700 border-slate-300 hover:bg-slate-200"
                            : "bg-emerald-100 text-emerald-900 border-emerald-300 hover:bg-emerald-200"
                        }`}
                        title={isTe ? "స్టేటస్ మార్చడానికి క్లిక్ చేయండి (Toggle Open/Closed)" : "Click to toggle Open/Closed status"}
                      >
                        <span className={`w-1.5 h-1.5 rounded-full ${isClosed ? "bg-slate-400" : "bg-emerald-600"}`} />
                        <span>{isClosed ? "Closed" : "Open"}</span>
                      </button>
                    </td>

                    {/* MIDDLE COLUMNS */}
                    {viewMode === "5col" ? (
                      /* 5 CONSOLIDATED MASTER COLUMNS (AS REQUESTED) WITH SEPARATE CELLS/INPUTS */
                      <>
                        {/* 1. Job Details & Dates (Box 1) - EDITABLE */}
                        <td className={`${cellPadding} border-r border-slate-200 align-top min-w-[240px]`}>
                          <div className="space-y-1.5 p-1 bg-blue-50/40 rounded-lg border border-blue-200/60 text-left">
                            {/* JC No */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-blue-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-blue-900 w-16 shrink-0">{isTe ? "JC నెం:" : "JC No:"}</span>
                              <input
                                type="text"
                                value={getRowValue(card, "jobNo")}
                                onChange={(e) => handleRowFieldChange(card.id, "jobNo", e.target.value)}
                                className="text-[11px] font-mono font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Complaint Date */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-blue-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-blue-900 w-16 shrink-0">{isTe ? "కంప్లైంట్:" : "Complaint:"}</span>
                              <input
                                type="date"
                                value={toInputDate(getRowValue(card, "complaintDate"))}
                                onChange={(e) => handleRowFieldChange(card.id, "complaintDate", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Open Date */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-blue-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-blue-900 w-16 shrink-0">{isTe ? "ఓపెన్ తేదీ:" : "Open Date:"}</span>
                              <input
                                type="date"
                                value={toInputDate(getRowValue(card, "jobDate"))}
                                onChange={(e) => handleRowFieldChange(card.id, "jobDate", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Branch */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-blue-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-blue-900 w-16 shrink-0">{isTe ? "బ్రాంచ్:" : "Branch:"}</span>
                              <select
                                value={getRowValue(card, "branch")}
                                onChange={(e) => handleRowFieldChange(card.id, "branch", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full cursor-pointer"
                              >
                                <option value="">-- select --</option>
                                {branchesList.map((b) => (
                                  <option key={b} value={b}>{b}</option>
                                ))}
                              </select>
                            </div>

                            {/* File No */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-blue-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-blue-900 w-16 shrink-0">{isTe ? "ఫైల్ నెం:" : "File No:"}</span>
                              <input
                                type="text"
                                value={getRowValue(card, "historyFileNo")}
                                onChange={(e) => handleRowFieldChange(card.id, "historyFileNo", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>
                          </div>
                        </td>

                        {/* 2. Customer Details (Box 2) - DISPLAY ONLY */}
                        <td className={`${cellPadding} border-r border-slate-200 align-top min-w-[240px]`}>
                          <div className="space-y-1.5 p-1 bg-slate-50 rounded-lg border border-slate-200/80 text-left">
                            {/* Customer Name */}
                            <div className="bg-white p-1 rounded border border-slate-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-slate-400 uppercase tracking-wider">{isTe ? "కస్టమర్ పేరు" : "Customer Name"}</span>
                              <span className="text-[11px] font-extrabold text-slate-900 truncate" title={card.custName || "—"}>
                                👤 {card.custName || "—"}
                              </span>
                            </div>

                            {/* Father Name */}
                            <div className="bg-white p-1 rounded border border-slate-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-slate-400 uppercase tracking-wider">{isTe ? "తండ్రి పేరు" : "Father's Name"}</span>
                              <span className="text-[11px] font-semibold text-slate-700 truncate" title={card.fatherName || "—"}>
                                S/O: {card.fatherName || "—"}
                              </span>
                            </div>

                            {/* Village */}
                            <div className="bg-white p-1 rounded border border-slate-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-slate-400 uppercase tracking-wider">{isTe ? "గ్రామం" : "Village"}</span>
                              <span className="text-[11px] font-semibold text-slate-700 truncate" title={card.village || "—"}>
                                📍 {card.village || "—"}
                              </span>
                            </div>

                            {/* Mandal */}
                            <div className="bg-white p-1 rounded border border-slate-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-slate-400 uppercase tracking-wider">{isTe ? "మండలం" : "Mandal"}</span>
                              <span className="text-[11px] font-semibold text-slate-700 truncate" title={card.mandal || "—"}>
                                🌾 {card.mandal || "—"}
                              </span>
                            </div>

                            {/* Phone No */}
                            <div className="bg-white p-1 rounded border border-slate-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-slate-400 uppercase tracking-wider">{isTe ? "మొబైల్ ఫోన్" : "Phone No"}</span>
                              <span className="text-[11px] font-mono font-bold text-emerald-800">
                                📞 {card.ownerMob || card.phNo || "—"}
                              </span>
                            </div>
                          </div>
                        </td>

                        {/* 3. Tractor Details (Box 3) - DISPLAY ONLY */}
                        <td className={`${cellPadding} border-r border-slate-200 align-top min-w-[240px]`}>
                          <div className="space-y-1.5 p-1 bg-amber-50/40 rounded-lg border border-amber-200/60 text-left">
                            {/* Tractor Model */}
                            <div className="bg-white p-1 rounded border border-amber-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-amber-800 uppercase tracking-wider">{isTe ? "ట్రాక్టర్ మోడల్" : "Tractor Model"}</span>
                              <span className="text-[11px] font-extrabold text-slate-900 truncate">
                                🚜 {card.model || "—"}
                              </span>
                            </div>

                            {/* Model Type */}
                            <div className="bg-white p-1 rounded border border-amber-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-amber-800 uppercase tracking-wider">{isTe ? "మోడల్ రకం" : "Model Type"}</span>
                              <span className="text-[11px] font-bold text-slate-700 truncate">
                                ⚙️ {card.modelType || "—"}
                              </span>
                            </div>

                            {/* Chassis No */}
                            <div className="bg-white p-1 rounded border border-amber-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-amber-800 uppercase tracking-wider">{isTe ? "ఛాసిస్ నెం." : "Chassis No"}</span>
                              <span className="text-[11px] font-mono font-bold text-indigo-950 truncate">
                                🆔 {card.chassisNo || "—"}
                              </span>
                            </div>

                            {/* Engine No */}
                            <div className="bg-white p-1 rounded border border-amber-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-amber-800 uppercase tracking-wider">{isTe ? "ఇంజన్ నెం." : "Engine No"}</span>
                              <span className="text-[11px] font-mono font-semibold text-slate-700 truncate">
                                🔌 {card.engineNo || "—"}
                              </span>
                            </div>

                            {/* Delivery Date */}
                            <div className="bg-white p-1 rounded border border-amber-100 shadow-2xs flex flex-col">
                              <span className="text-[8.5px] font-bold text-amber-800 uppercase tracking-wider">{isTe ? "డెలివరీ తేదీ" : "Delivery Date"}</span>
                              <span className="text-[11px] font-semibold text-slate-700">
                                📅 {formatDate(card.dateOfDelivery || card.installDate)}
                              </span>
                            </div>
                          </div>
                        </td>

                        {/* 4. Service & Repairs (Box 4) - EDITABLE */}
                        <td className={`${cellPadding} border-r border-slate-200 align-top min-w-[250px]`}>
                          <div className="space-y-1.5 p-1 bg-emerald-50/40 rounded-lg border border-emerald-200/60 text-left">
                            {/* Hour Meter */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-emerald-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-emerald-900 w-16 shrink-0">{isTe ? "గంటలు:" : "Hrs Run:"}</span>
                              <input
                                type="number"
                                value={getRowValue(card, "hourMeter") !== undefined ? getRowValue(card, "hourMeter") : (card.hourMeter || card.hrsRun || "")}
                                onChange={(e) => handleRowFieldChange(card.id, "hourMeter", e.target.value)}
                                className="text-[11px] font-mono font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Service Type */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-emerald-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-emerald-900 w-16 shrink-0">{isTe ? "సర్వీస్ రకం:" : "Service Type:"}</span>
                              <select
                                value={getRowValue(card, "serviceType")}
                                onChange={(e) => handleRowFieldChange(card.id, "serviceType", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full cursor-pointer"
                              >
                                <option value="">-- select --</option>
                                {serviceTypes.map((st) => (
                                  <option key={st} value={st}>{st}</option>
                                ))}
                              </select>
                            </div>

                            {/* Free Service List */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-emerald-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-emerald-900 w-16 shrink-0">{isTe ? "ఫ్రీ సర్వీస్:" : "Free Serv:"}</span>
                              <input
                                type="text"
                                value={getRowValue(card, "freeServiceList")}
                                onChange={(e) => handleRowFieldChange(card.id, "freeServiceList", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Extra Repairs */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-emerald-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-emerald-900 w-16 shrink-0">{isTe ? "రిపేర్లు:" : "Repairs:"}</span>
                              <input
                                type="text"
                                value={getRowValue(card, "extraRepairs") !== undefined ? getRowValue(card, "extraRepairs") : (card.extraRepairs || card.problemDescription || "")}
                                onChange={(e) => handleRowFieldChange(card.id, "extraRepairs", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Grand Total */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-emerald-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-emerald-900 w-16 shrink-0">{isTe ? "బిల్ (₹):" : "Bill (₹):"}</span>
                              <input
                                type="number"
                                value={getRowValue(card, "gTotal")}
                                onChange={(e) => handleRowFieldChange(card.id, "gTotal", e.target.value)}
                                className="text-[11px] font-bold text-emerald-950 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>
                          </div>
                        </td>

                        {/* 5. Closure & Technician (Box 5) - EDITABLE */}
                        <td className={`${cellPadding} border-r border-slate-200 align-top min-w-[260px]`}>
                          <div className="space-y-1.5 p-1 bg-purple-50/40 rounded-lg border border-purple-200/60 text-left">
                            {/* Actual Closed Date */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-purple-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-purple-900 w-16 shrink-0">{isTe ? "క్లోజ్ తేదీ:" : "Closed Date:"}</span>
                              <input
                                type="date"
                                value={toInputDate(getRowValue(card, "actualClosedDate") !== undefined ? getRowValue(card, "actualClosedDate") : (card.actualClosedDate || card.dateTimeOut))}
                                onChange={(e) => handleRowFieldChange(card.id, "actualClosedDate", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Mechanic */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-purple-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-purple-900 w-16 shrink-0">{isTe ? "టెక్నీషియన్:" : "Technician:"}</span>
                              <input
                                type="text"
                                list={`excel-mech-list-${card.id}`}
                                value={getRowValue(card, "mechanic")}
                                onChange={(e) => handleRowFieldChange(card.id, "mechanic", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                              <datalist id={`excel-mech-list-${card.id}`}>
                                {Array.from(new Set(mechanicsList || [])).map((m, idx) => (
                                  <option key={`mech-opt-${card.id}-${m}-${idx}`} value={m} />
                                ))}
                              </datalist>
                            </div>

                            {/* Service Location */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-purple-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-purple-900 w-16 shrink-0">{isTe ? "స్థలం:" : "Location:"}</span>
                              <select
                                value={getRowValue(card, "serviceLocation")}
                                onChange={(e) => handleRowFieldChange(card.id, "serviceLocation", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full cursor-pointer"
                              >
                                <option value="workshop">🏭 {isTe ? "వర్క్‌షాప్" : "Workshop"}</option>
                                <option value="doorstep">🚜 {isTe ? "రైతు ఇంటి వద్ద" : "Doorstep"}</option>
                                <option value="field">🌾 {isTe ? "ఫీల్డ్" : "Field"}</option>
                                <option value="camp">⛺ {isTe ? "క్యాంప్" : "Camp"}</option>
                              </select>
                            </div>

                            {/* Bill No */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-purple-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-purple-900 w-16 shrink-0">{isTe ? "బిల్ నెం:" : "Bill No:"}</span>
                              <input
                                type="text"
                                value={getRowValue(card, "billNo")}
                                onChange={(e) => handleRowFieldChange(card.id, "billNo", e.target.value)}
                                className="text-[11px] font-mono font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Reasons for Analysis / Problem Remarks */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-purple-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-purple-900 w-16 shrink-0">{isTe ? "రిమార్క్స్:" : "Remarks:"}</span>
                              <input
                                type="text"
                                value={getRowValue(card, "reasonsForAnalysis") !== undefined ? getRowValue(card, "reasonsForAnalysis") : (card.reasonsForAnalysis || card.problemDescription || "")}
                                onChange={(e) => handleRowFieldChange(card.id, "reasonsForAnalysis", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>

                            {/* Telecalling */}
                            <div className="flex items-center gap-1 bg-white p-1 rounded border border-purple-100 shadow-2xs">
                              <span className="text-[9px] font-extrabold text-purple-900 w-16 shrink-0">{isTe ? "టెలికాల్:" : "Telecall:"}</span>
                              <input
                                type="text"
                                value={getRowValue(card, "telecalling")}
                                onChange={(e) => handleRowFieldChange(card.id, "telecalling", e.target.value)}
                                className="text-[11px] font-bold text-slate-800 bg-transparent border-0 p-0 focus:ring-0 focus:outline-none w-full"
                              />
                            </div>
                          </div>
                        </td>
                      </>
                    ) : (
                      /* DETAILED 26 EXCEL COLUMNS */
                      detailedColumns.map((col) => {
                        const val = getCardColValue(card, col.key);
                        return (
                          <td
                            key={col.key}
                            className={`${cellPadding} border-r border-slate-200 truncate max-w-xs font-medium`}
                            title={val}
                          >
                            {col.key === "jobNo" || col.key === "onlineJobCardNo" || col.key === "chassisNo" ? (
                              <span className="font-mono font-bold text-slate-900">{val || "—"}</span>
                            ) : col.key === "custName" || col.key === "model" ? (
                              <span className="font-bold text-slate-900">{val || "—"}</span>
                            ) : (
                              val || "—"
                            )}
                          </td>
                        );
                      })
                    )}

                    {/* LAST COLUMN: ACTIONS (EYE, EDIT, PRINT, SAVE OPTIONS) */}
                    <td className={`${cellPadding} text-center sticky right-0 z-20 bg-white shadow-[-4px_0_6px_-2px_rgba(0,0,0,0.06)]`}>
                      <div className="flex items-center justify-center gap-1">
                        {/* 1. EYE SYMBOL: VIEW JOB CARD */}
                        <button
                          type="button"
                          onClick={() => onView(card)}
                          className="p-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-md transition-all cursor-pointer shadow-2xs hover:scale-110 active:scale-95"
                          title={isTe ? "జాబ్ కార్డ్ చూడటానికి (View Job Card)" : "View Job Card Details"}
                        >
                          <Eye className="w-3.5 h-3.5" />
                        </button>

                        {/* 2. EDIT SYMBOL: EDIT JOB CARD */}
                        <button
                          type="button"
                          onClick={() => onEdit(card)}
                          className="p-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 rounded-md transition-all cursor-pointer shadow-2xs hover:scale-110 active:scale-95"
                          title={isTe ? "జాబ్ కార్డ్ ఎడిట్ చేయి (Edit Job Card)" : "Edit Job Card"}
                        >
                          <PenLine className="w-3.5 h-3.5" />
                        </button>

                        {/* 3. PRINT SYMBOL: PRINT JOB CARD */}
                        <button
                          type="button"
                          onClick={() => onPrint(card)}
                          className="p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md transition-all cursor-pointer shadow-2xs hover:scale-110 active:scale-95"
                          title={isTe ? "ప్రింట్ తీయండి (Print Job Card)" : "Print Job Card"}
                        >
                          <Printer className="w-3.5 h-3.5" />
                        </button>

                        {/* 4. SAVE OPTION: QUICK SAVE / CONFIRM */}
                        <button
                          type="button"
                          onClick={() => handleSaveRow(card)}
                          disabled={isSaving}
                          className={`p-1.5 rounded-md transition-all cursor-pointer shadow-2xs hover:scale-110 active:scale-95 ${
                            isSavedSuccess
                              ? "bg-emerald-600 text-white font-bold"
                              : "bg-emerald-50 hover:bg-emerald-100 text-emerald-800"
                          }`}
                          title={isTe ? "మార్పులను సేవ్ చేయి (Save to DB)" : "Save to Database"}
                        >
                          {isSavedSuccess ? (
                            <Check className="w-3.5 h-3.5 text-white animate-bounce" />
                          ) : (
                            <Save className="w-3.5 h-3.5" />
                          )}
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* SPREADSHEET BOTTOM STATUS BAR & PAGINATION */}
      <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-50 border border-slate-200 p-2.5 rounded-lg text-xs">
        <div className="flex items-center gap-2 text-slate-600 font-semibold text-[11px]">
          <span>
            {isTe ? "చూపిస్తున్నవి:" : "Showing"} <strong className="text-slate-900">{processedCards.length === 0 ? 0 : (currentPage - 1) * pageSize + 1}</strong> - <strong className="text-slate-900">{Math.min(currentPage * pageSize, processedCards.length)}</strong> of <strong className="text-slate-900">{processedCards.length}</strong>
          </span>
          {selectedIds.length > 0 && (
            <span className="px-2 py-0.5 bg-emerald-100 text-emerald-900 rounded font-bold">
              {selectedIds.length} {isTe ? "ఎంపిక చేయబడ్డాయి" : "selected"}
            </span>
          )}
        </div>

        <div className="flex items-center gap-2">
          <div className="flex items-center gap-1">
            <span className="text-[10px] font-bold text-slate-600">
              {isTe ? "పేజీకి:" : "Per page:"}
            </span>
            <select
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setCurrentPage(1);
              }}
              className="text-[10.5px] font-bold py-0.5 px-1 bg-white border border-slate-300 rounded outline-none cursor-pointer"
            >
              <option value={10}>10</option>
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={500}>500</option>
            </select>
          </div>

          <div className="flex items-center gap-1">
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => setCurrentPage(1)}
              className="px-2 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
            >
              « First
            </button>
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              className="px-2 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
            >
              ‹ Prev
            </button>
            <span className="px-2.5 py-0.5 text-[10px] font-bold text-slate-800 bg-white border border-slate-300 rounded shadow-2xs">
              {currentPage} / {totalPages}
            </span>
            <button
              type="button"
              disabled={currentPage >= totalPages}
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              className="px-2 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
            >
              Next ›
            </button>
            <button
              type="button"
              disabled={currentPage >= totalPages}
              onClick={() => setCurrentPage(totalPages)}
              className="px-2 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
            >
              Last »
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
