import React, { useState, useMemo, useEffect } from 'react';
import {
  MapPin,
  Navigation,
  Calendar,
  Plus,
  ChevronDown,
  ChevronRight,
  Phone,
  MessageSquare,
  Wrench,
  Users,
  CheckCircle,
  Clock,
  Search,
  Download,
  X,
  FileText,
  Building2,
  Truck,
  ExternalLink,
  Map as MapIcon,
  CalendarDays,
  Check,
  Filter,
  Folder,
  FolderOpen,
  List,
  GitFork,
  ArrowUpDown,
  Sparkles,
  Layers,
  ChevronUp
} from 'lucide-react';
import {
  DEALERSHIP_DATA,
  normalizeGeoStr,
  getGoogleMapsDirectionsUrl,
  DealershipInfo,
  BranchInfo,
  MandalInfo,
  VillageInfo
} from '../data/dealershipData';
import { Language } from '../translations';
import { apiSaveServiceCamp, apiDeleteServiceCamp } from '../api';
import { VillageCustomerInline } from './VillageCustomerInline';
import { BroadcastQueueModal } from './BroadcastQueueModal';
import { KrishnaDistrictRouteMap } from './KrishnaDistrictRouteMap';
import { masterGeoDirectory } from '../lib/geoMatcher';

export interface ServiceCamp {
  id: string;
  dealershipCode: '4731' | '4732' | string;
  branch: string;
  mandal: string;
  village: string;
  campDate: string;
  targetTractors: string;
  supervisor: string;
  mechanic: string;
  status: 'Upcoming' | 'In Progress' | 'Completed' | 'Cancelled';
  serviceTypeExpected: string;
  offers: string;
  contactPerson: string;
  contactPhone: string;
  notes: string;
  attendedCount: string;
  venue?: string;
  createdAt: string;
}

interface ServiceCampPlanningProps {
  language: Language;
  customers: any[];
  jobCards: any[];
  complaints: any[];
  staffList: any[];
  serviceCamps: ServiceCamp[];
  onUpdateServiceCamps: (camps: ServiceCamp[]) => void;
  onNavigateToJobCard?: (customerId?: string) => void;
}

export const ServiceCampPlanning: React.FC<ServiceCampPlanningProps> = ({
  language,
  customers = [],
  jobCards = [],
  complaints = [],
  staffList = [],
  serviceCamps = [],
  onUpdateServiceCamps
}) => {
  const isTe = language === 'te';

  // Navigation main tab: 'planning' (Villages Tree/Table), 'camps_list' (Scheduled Camps), or 'district_map' (Krishna & NTR District Route Map)
  const [activeTab, setActiveTab] = useState<'planning' | 'camps_list' | 'district_map'>('planning');

  // View style inside planning: compact tree vs flat table
  const [viewStyle, setViewStyle] = useState<'tree' | 'table'>('tree');

  // Dropdown filter selections ("Adda Box" / Horizontal Bar Controls)
  const [selectedHub, setSelectedHub] = useState<'all' | '4731' | '4732'>('all');
  const [selectedBranch, setSelectedBranch] = useState<string>('all');
  const [selectedMandal, setSelectedMandal] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Table pagination
  const [tablePage, setTablePage] = useState(1);
  const [pageSize, setPageSize] = useState(25);

  // Expanded nodes for Tree view (compact directory tree)
  const [expandedHubs, setExpandedHubs] = useState<Record<string, boolean>>({
    '4731': true,
    '4732': true
  });
  const [expandedBranches, setExpandedBranches] = useState<Record<string, boolean>>({
    tiruvuru: true,
    nuzvidu: false,
    nandigama: false,
    gudiwada: true,
    poranki_vijayawada: false,
    machilipatnam_21: false
  });
  const [expandedMandals, setExpandedMandals] = useState<Record<string, boolean>>({
    'tiruvuru-Tiruvuru': true,
    'gudiwada-Gudivada': true
  });

  // Expanded Villages for Inline Customer Intelligence
  const [expandedVillages, setExpandedVillages] = useState<Record<string, boolean>>({});

  // Dedicated Large Customer Data Box ("Pedda Box") State
  const [selectedVillageKey, setSelectedVillageKey] = useState<string>('');
  const [activeVillageDetails, setActiveVillageDetails] = useState<{
    villageName: string;
    villageTelugu?: string;
    mandalName: string;
    mandalTelugu?: string;
    branchName: string;
    hubCode: string;
    distanceKm: number;
    approxTravelTime?: string;
  } | null>(null);

  const handleSelectVillage = (
    village: VillageInfo,
    mandal: MandalInfo,
    branch: BranchInfo,
    hubCode: string
  ) => {
    const vKey = `${branch.id}-${mandal.name}-${village.name}`;
    setSelectedVillageKey(vKey);
    setActiveVillageDetails({
      villageName: village.name,
      villageTelugu: village.teluguName,
      mandalName: mandal.name,
      mandalTelugu: mandal.teluguName,
      branchName: branch.name,
      hubCode: String(hubCode),
      distanceKm: village.distanceKm,
      approxTravelTime: village.approxTravelTime
    });
  };

  // WhatsApp & Contact tracking
  const [contactedCustomers, setContactedCustomers] = useState<Record<string, boolean>>({});

  // Active Broadcast Assistant modal data
  const [broadcastData, setBroadcastData] = useState<{
    villageName: string;
    mandalName: string;
    branchName: string;
    customers: any[];
    messageTemplate: string;
  } | null>(null);

  // Toast notification
  const [toastMsg, setToastMsg] = useState<string | null>(null);

  const showToast = (msg: string) => {
    setToastMsg(msg);
    setTimeout(() => {
      setToastMsg(null);
    }, 3500);
  };

  const toggleVillage = (villageKey: string) => {
    setExpandedVillages(prev => ({
      ...prev,
      [villageKey]: !prev[villageKey]
    }));
  };

  const handleMarkContacted = (customerKey: string) => {
    setContactedCustomers(prev => ({
      ...prev,
      [customerKey]: true
    }));
  };

  // Modal for Planning a New Camp or Editing existing
  const [isCampModalOpen, setIsCampModalOpen] = useState(false);
  const [editingCamp, setEditingCamp] = useState<Partial<ServiceCamp> | null>(null);
  const [savingCamp, setSavingCamp] = useState(false);

  // Filter for Camps List Tab
  const [campStatusFilter, setCampStatusFilter] = useState<'all' | 'Upcoming' | 'In Progress' | 'Completed' | 'Cancelled'>('all');

  // Expand / Collapse toggles
  const toggleHub = (hubCode: string) => {
    setExpandedHubs(prev => ({ ...prev, [hubCode]: !prev[hubCode] }));
  };

  const toggleBranch = (branchId: string) => {
    setExpandedBranches(prev => ({ ...prev, [branchId]: !prev[branchId] }));
  };

  const toggleMandal = (mandalKey: string) => {
    setExpandedMandals(prev => ({ ...prev, [mandalKey]: !prev[mandalKey] }));
  };

  const expandAll = () => {
    const allH: Record<string, boolean> = { '4731': true, '4732': true };
    const allB: Record<string, boolean> = {};
    const allM: Record<string, boolean> = {};
    ['4731', '4732'].forEach(code => {
      const hub = DEALERSHIP_DATA[code as '4731' | '4732'];
      hub.branches.forEach(b => {
        allB[b.id] = true;
        b.mandals.forEach(m => {
          allM[`${b.id}-${m.name}`] = true;
        });
      });
    });
    setExpandedHubs(allH);
    setExpandedBranches(allB);
    setExpandedMandals(allM);
  };

  const collapseAll = () => {
    setExpandedHubs({});
    setExpandedBranches({});
    setExpandedMandals({});
    setExpandedVillages({});
  };

  // Build list of branches according to selected hub
  const availableBranches = useMemo(() => {
    const branches: BranchInfo[] = [];
    if (selectedHub === 'all' || selectedHub === '4731') {
      branches.push(...DEALERSHIP_DATA['4731'].branches);
    }
    if (selectedHub === 'all' || selectedHub === '4732') {
      branches.push(...DEALERSHIP_DATA['4732'].branches);
    }
    return branches;
  }, [selectedHub]);

  // Build list of mandals according to selected branch & hub
  const availableMandals = useMemo(() => {
    if (selectedBranch !== 'all') {
      const branch = availableBranches.find(b => b.name === selectedBranch || b.id === selectedBranch);
      return branch ? branch.mandals.map(m => m.name) : [];
    }
    const set = new Set<string>();
    availableBranches.forEach(b => {
      b.mandals.forEach(m => set.add(m.name));
    });
    return Array.from(set).sort();
  }, [selectedBranch, availableBranches]);

  // Handle Hub change from dropdown
  const handleHubChange = (hub: 'all' | '4731' | '4732') => {
    setSelectedHub(hub);
    setSelectedBranch('all');
    setSelectedMandal('all');
    setTablePage(1);
  };

  // Handle Branch change from dropdown
  const handleBranchChange = (branch: string) => {
    setSelectedBranch(branch);
    setSelectedMandal('all');
    setTablePage(1);
    if (branch !== 'all') {
      const bObj = availableBranches.find(b => b.name === branch || b.id === branch);
      if (bObj) {
        setExpandedBranches(prev => ({ ...prev, [bObj.id]: true }));
      }
    }
  };

  // Customer Village Mapping Index (AI Smart Matcher with Master Geo Directory & Google Maps)
  const customersByLocation = useMemo(() => {
    const map: Record<string, any[]> = {};

    customers.forEach(c => {
      const rawV = c.village || c.custAddr || '';
      const rawM = c.mandal || '';
      const rawAddr = c.custAddr || c.address || '';

      // Match using Google Maps & Dealership master directory engine
      const match = masterGeoDirectory.matchLocation(rawV, rawM, rawAddr);

      const isAiMatched = match.matched && (
        normalizeGeoStr(match.villageName) !== normalizeGeoStr(rawV) ||
        (Boolean(rawM) && normalizeGeoStr(match.mandalName) !== normalizeGeoStr(rawM))
      );

      const enrichedCust = {
        ...c,
        _geoMatch: match,
        _canonicalVillage: match.villageName,
        _canonicalMandal: match.mandalName,
        _canonicalBranch: match.branchName,
        _dealershipCode: match.dealershipCode,
        _distanceKm: match.distanceKm,
        _isAiMatched: isAiMatched,
        _matchConfidence: match.confidence,
        _matchType: match.matchType,
        _originalVillage: rawV
      };

      // 1. Index under canonical village key
      const canonicalVKey = normalizeGeoStr(match.villageName);
      if (canonicalVKey) {
        if (!map[canonicalVKey]) map[canonicalVKey] = [];
        map[canonicalVKey].push(enrichedCust);
      }

      // 2. Also index under raw village key if distinct
      const rawVKey = normalizeGeoStr(rawV);
      if (rawVKey && rawVKey !== canonicalVKey) {
        if (!map[rawVKey]) map[rawVKey] = [];
        map[rawVKey].push(enrichedCust);
      }

      // 3. Index under canonical mandal key
      const canonicalMKey = `mandal_${normalizeGeoStr(match.mandalName)}`;
      if (!map[canonicalMKey]) map[canonicalMKey] = [];
      map[canonicalMKey].push(enrichedCust);

      // 4. Also index under raw mandal key
      const rawMKey = `mandal_${normalizeGeoStr(rawM)}`;
      if (rawMKey && rawMKey !== canonicalMKey) {
        if (!map[rawMKey]) map[rawMKey] = [];
        map[rawMKey].push(enrichedCust);
      }
    });

    return map;
  }, [customers]);

  // Helper to get customers for a village
  const getCustomersForVillage = (villageName: string, mandalName?: string): any[] => {
    const vKey = normalizeGeoStr(villageName);
    let list = customersByLocation[vKey] || [];

    if (list.length === 0 && villageName.length > 3) {
      const sub = vKey.substring(0, Math.min(vKey.length, 6));
      for (const [k, vCusts] of Object.entries(customersByLocation)) {
        if (!k.startsWith('mandal_') && (k.includes(sub) || sub.includes(k))) {
          list = list.concat(vCusts);
        }
      }
    }

    // Deduplicate by chassisKey or chassisNo or id or mobile
    const uniqueMap = new Map();
    list.forEach(c => {
      const key = c.chassisKey || c.chassisNo || c.id || `${c.custName}_${c.ownerMob || c.phone}`;
      if (key && !uniqueMap.has(key)) {
        uniqueMap.set(key, c);
      }
    });

    return Array.from(uniqueMap.values());
  };

  // Helper to count tractors in a village
  const getTractorCountForVillage = (villageName: string, mandalName?: string): number => {
    return getCustomersForVillage(villageName, mandalName).length;
  };

  // Quick Stats
  const stats = useMemo(() => {
    let totalTractors4731 = 0;
    let totalTractors4732 = 0;
    let totalVillages4731 = 0;
    let totalVillages4732 = 0;

    DEALERSHIP_DATA['4731'].branches.forEach(b => {
      b.mandals.forEach(m => {
        totalVillages4731 += m.villages.length;
        m.villages.forEach(v => {
          totalTractors4731 += getTractorCountForVillage(v.name, m.name);
        });
      });
    });

    DEALERSHIP_DATA['4732'].branches.forEach(b => {
      b.mandals.forEach(m => {
        totalVillages4732 += m.villages.length;
        m.villages.forEach(v => {
          totalTractors4732 += getTractorCountForVillage(v.name, m.name);
        });
      });
    });

    const upcomingCamps = serviceCamps.filter(c => c.status === 'Upcoming').length;
    const completedCamps = serviceCamps.filter(c => c.status === 'Completed').length;
    const inProgressCamps = serviceCamps.filter(c => c.status === 'In Progress').length;

    return {
      totalCustomers: customers.length,
      totalTractors4731,
      totalTractors4732,
      totalVillages: totalVillages4731 + totalVillages4732,
      totalPlannedCamps: serviceCamps.length,
      upcomingCamps,
      completedCamps,
      inProgressCamps
    };
  }, [customers, serviceCamps, customersByLocation]);

  // Filtered Dealership Hierarchy for Tree & Table
  const filteredDealerships = useMemo(() => {
    const list: DealershipInfo[] = [];
    const q = searchQuery.trim().toLowerCase();

    (['4731', '4732'] as const).forEach(code => {
      if (selectedHub !== 'all' && selectedHub !== code) return;

      const hub = DEALERSHIP_DATA[code];
      const filteredBranches: BranchInfo[] = [];

      hub.branches.forEach(b => {
        if (selectedBranch !== 'all' && selectedBranch !== b.name && selectedBranch !== b.id) {
          return;
        }

        const filteredMandals: MandalInfo[] = [];

        b.mandals.forEach(m => {
          if (selectedMandal !== 'all' && selectedMandal !== m.name) {
            return;
          }

          if (!q) {
            filteredMandals.push(m);
            return;
          }

          const branchMatch = b.name.toLowerCase().includes(q) || b.teluguName.toLowerCase().includes(q);
          const mandalMatch = m.name.toLowerCase().includes(q) || (m.teluguName && m.teluguName.toLowerCase().includes(q));
          const filteredVillages = m.villages.filter(v =>
            v.name.toLowerCase().includes(q) || (v.teluguName && v.teluguName.toLowerCase().includes(q))
          );

          if (branchMatch || mandalMatch || filteredVillages.length > 0) {
            filteredMandals.push({
              ...m,
              villages: branchMatch || mandalMatch ? m.villages : filteredVillages
            });
          }
        });

        if (filteredMandals.length > 0) {
          filteredBranches.push({
            ...b,
            mandals: filteredMandals
          });
        }
      });

      if (filteredBranches.length > 0) {
        list.push({
          ...hub,
          branches: filteredBranches
        });
      }
    });

    return list;
  }, [selectedHub, selectedBranch, selectedMandal, searchQuery]);

  // Auto-select first village with tractors for the Pedda Box so customer data is ready immediately
  useEffect(() => {
    if (!activeVillageDetails && filteredDealerships.length > 0) {
      for (const hub of filteredDealerships) {
        for (const b of hub.branches) {
          for (const m of b.mandals) {
            for (const v of m.villages) {
              const count = getCustomersForVillage(v.name, m.name).length;
              if (count > 0) {
                handleSelectVillage(v, m, b, hub.code);
                return;
              }
            }
          }
        }
      }
      // Fallback to the very first village if none have customer tractors
      const firstHub = filteredDealerships[0];
      const firstB = firstHub?.branches[0];
      const firstM = firstB?.mandals[0];
      const firstV = firstM?.villages[0];
      if (firstV && firstM && firstB && firstHub) {
        handleSelectVillage(firstV, firstM, firstB, firstHub.code);
      }
    }
  }, [filteredDealerships, activeVillageDetails, customersByLocation]);

  // Flat list of villages for Table View
  const flatVillagesList = useMemo(() => {
    const list: {
      hubCode: '4731' | '4732';
      branchName: string;
      branchTelugu: string;
      mandalName: string;
      mandalTelugu?: string;
      villageName: string;
      villageTelugu?: string;
      distanceKm: number;
      approxTravelTime?: string;
      tractorCount: number;
      directionsUrl: string;
    }[] = [];

    filteredDealerships.forEach(hub => {
      hub.branches.forEach(branch => {
        branch.mandals.forEach(mandal => {
          mandal.villages.forEach(v => {
            const count = getTractorCountForVillage(v.name, mandal.name);
            const url = getGoogleMapsDirectionsUrl(branch.name, v.name, mandal.name);
            list.push({
              hubCode: hub.code,
              branchName: branch.name,
              branchTelugu: branch.teluguName,
              mandalName: mandal.name,
              mandalTelugu: mandal.teluguName,
              villageName: v.name,
              villageTelugu: v.teluguName,
              distanceKm: v.distanceKm,
              approxTravelTime: v.approxTravelTime,
              tractorCount: count,
              directionsUrl: url
            });
          });
        });
      });
    });

    return list;
  }, [filteredDealerships, customersByLocation]);

  // Paginated Flat Villages for Table View
  const paginatedVillages = useMemo(() => {
    const start = (tablePage - 1) * pageSize;
    return flatVillagesList.slice(start, start + pageSize);
  }, [flatVillagesList, tablePage, pageSize]);

  const totalTablePages = Math.ceil(flatVillagesList.length / pageSize) || 1;

  // Filtered Service Camps for Scheduled Camps View
  const filteredCamps = useMemo(() => {
    return serviceCamps.filter(c => {
      if (campStatusFilter !== 'all' && c.status !== campStatusFilter) return false;
      if (selectedHub !== 'all' && c.dealershipCode !== selectedHub) return false;
      if (selectedBranch !== 'all' && c.branch !== selectedBranch) return false;
      if (selectedMandal !== 'all' && c.mandal !== selectedMandal) return false;
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        return (
          c.village.toLowerCase().includes(q) ||
          c.mandal.toLowerCase().includes(q) ||
          c.branch.toLowerCase().includes(q) ||
          c.supervisor.toLowerCase().includes(q) ||
          c.mechanic.toLowerCase().includes(q) ||
          c.contactPerson.toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [serviceCamps, campStatusFilter, selectedHub, selectedBranch, selectedMandal, searchQuery]);

  // Handle Camp Plan Form Open
  const handleOpenNewCampModal = (prefill?: {
    dealershipCode?: '4731' | '4732' | string;
    branch?: string;
    mandal?: string;
    village?: string;
  }) => {
    const code = prefill?.dealershipCode || (selectedHub !== 'all' ? selectedHub : '4731');
    const branch = prefill?.branch || (selectedBranch !== 'all' ? selectedBranch : 'Tiruvuru');
    const hub = DEALERSHIP_DATA[code as '4731' | '4732'] || DEALERSHIP_DATA['4731'];
    const branchObj = hub.branches.find(b => b.name === branch) || hub.branches[0];
    const mandal = prefill?.mandal || (selectedMandal !== 'all' ? selectedMandal : branchObj.mandals[0]?.name || '');

    setEditingCamp({
      id: `camp_${Date.now()}`,
      dealershipCode: code,
      branch: branchObj.name,
      mandal: mandal,
      village: prefill?.village || '',
      campDate: new Date().toISOString().split('T')[0],
      targetTractors: '15',
      supervisor: staffList.find(s => s.role?.toLowerCase().includes('supervisor'))?.name || '',
      mechanic: staffList.find(s => s.role?.toLowerCase().includes('mechanic'))?.name || '',
      status: 'Upcoming',
      serviceTypeExpected: 'Free 50-Point Inspection & Oil Service Camp',
      offers: 'Free 50-Point General Checkup, 10% Discount on Eicher Genuine Oil & Filters, Free Greasing & Battery Health Check',
      contactPerson: '',
      contactPhone: '',
      venue: 'Rythu Bharosa Kendram (RBK) / Gram Panchayat',
      notes: '',
      attendedCount: '0',
      createdAt: new Date().toISOString()
    });
    setIsCampModalOpen(true);
  };

  const handleEditCamp = (camp: ServiceCamp) => {
    setEditingCamp({ ...camp });
    setIsCampModalOpen(true);
  };

  // Save Service Camp
  const handleSaveCamp = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingCamp || !editingCamp.village || !editingCamp.campDate) {
      alert(isTe ? 'దయచేసి గ్రామం మరియు క్యాంప్ తేదీని ఎంచుకోండి.' : 'Please select village and camp date.');
      return;
    }

    setSavingCamp(true);
    const campToSave: ServiceCamp = {
      id: editingCamp.id || `camp_${Date.now()}`,
      dealershipCode: editingCamp.dealershipCode || '4731',
      branch: editingCamp.branch || 'Tiruvuru',
      mandal: editingCamp.mandal || '',
      village: editingCamp.village || '',
      campDate: editingCamp.campDate || '',
      targetTractors: editingCamp.targetTractors || '10',
      supervisor: editingCamp.supervisor || '',
      mechanic: editingCamp.mechanic || '',
      status: (editingCamp.status as any) || 'Upcoming',
      serviceTypeExpected: editingCamp.serviceTypeExpected || '',
      offers: editingCamp.offers || '',
      contactPerson: editingCamp.contactPerson || '',
      contactPhone: editingCamp.contactPhone || '',
      venue: editingCamp.venue || '',
      notes: editingCamp.notes || '',
      attendedCount: editingCamp.attendedCount || '0',
      createdAt: editingCamp.createdAt || new Date().toISOString()
    };

    try {
      await apiSaveServiceCamp(campToSave);

      const existingIdx = serviceCamps.findIndex(c => c.id === campToSave.id);
      let updated: ServiceCamp[];
      if (existingIdx >= 0) {
        updated = [...serviceCamps];
        updated[existingIdx] = campToSave;
      } else {
        updated = [campToSave, ...serviceCamps];
      }
      onUpdateServiceCamps(updated);
      setIsCampModalOpen(false);
      setEditingCamp(null);
    } catch (err) {
      console.error('Error saving camp:', err);
    } finally {
      setSavingCamp(false);
    }
  };

  // Delete Service Camp
  const handleDeleteCamp = async (id: string) => {
    if (!confirm(isTe ? 'ఈ సర్వీస్ క్యాంప్ ప్లాన్‌ను తొలగించాలా?' : 'Are you sure you want to delete this service camp plan?')) {
      return;
    }
    try {
      await apiDeleteServiceCamp(id);
      const updated = serviceCamps.filter(c => c.id !== id);
      onUpdateServiceCamps(updated);
    } catch (err) {
      console.error('Error deleting camp:', err);
    }
  };

  // Export Village Customer List to CSV
  const exportVillageCustomerExcel = (village: string, mandal: string, branch: string, custList: any[]) => {
    const headers = [
      'Sl No',
      'Customer Name',
      'Father Name',
      'Owner Mobile',
      'Village',
      'Mandal',
      'Branch',
      'Tractor Model',
      'Chassis No',
      'Delivery Date',
      'HFN'
    ];

    const rows = custList.map((c, i) => [
      i + 1,
      `"${(c.custName || c.cust_name || c.customerName || '').replace(/"/g, '""')}"`,
      `"${(c.fatherName || c.father_name || '').replace(/"/g, '""')}"`,
      `"${c.ownerMob || c.owner_mob || c.mobileNumber || ''}"`,
      `"${village}"`,
      `"${mandal}"`,
      `"${branch}"`,
      `"${c.tractorModel || c.tractor_model || c.model || 'Eicher Tractor'}"`,
      `"${c.chassisNo || c.chassis_no || ''}"`,
      `"${c.dateOfDelivery || c.date_of_delivery || ''}"`,
      `"${c.historyFileNo || c.history_file_no || ''}"`
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,\uFEFF' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `Eicher_Customers_${village}_${mandal}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Export Planned Camps to CSV
  const exportServiceCampsExcel = () => {
    const headers = [
      'Camp ID',
      'Dealership Code',
      'Branch',
      'Mandal',
      'Village',
      'Camp Date',
      'Target Tractors',
      'Attended Count',
      'Status',
      'Supervisor',
      'Mechanic',
      'Venue',
      'Contact Person',
      'Contact Phone',
      'Special Offers'
    ];

    const rows = serviceCamps.map(c => [
      `"${c.id}"`,
      `"${c.dealershipCode}"`,
      `"${c.branch}"`,
      `"${c.mandal}"`,
      `"${c.village}"`,
      `"${c.campDate}"`,
      `"${c.targetTractors}"`,
      `"${c.attendedCount || 0}"`,
      `"${c.status}"`,
      `"${c.supervisor || ''}"`,
      `"${c.mechanic || ''}"`,
      `"${(c.venue || '').replace(/"/g, '""')}"`,
      `"${(c.contactPerson || '').replace(/"/g, '""')}"`,
      `"${c.contactPhone || ''}"`,
      `"${(c.offers || '').replace(/"/g, '""')}"`
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,\uFEFF' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `Eicher_Service_Camps_Plan_${new Date().toISOString().split('T')[0]}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="space-y-3 pb-8">
      {/* ========================================================================= */}
      {/* 1. SLIM HORIZONTAL METRIC RIBBON (No massive cards, ultra-compact) */}
      {/* ========================================================================= */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xs p-3 flex flex-wrap items-center justify-between gap-3 text-xs">
        <div className="flex items-center gap-2">
          <span className="p-1.5 bg-blue-900 text-white rounded-lg shadow-2xs">
            <Truck className="w-4 h-4" />
          </span>
          <div>
            <h1 className="text-sm md:text-base font-black text-slate-900 leading-tight">
              {isTe ? '⛺ సర్వీస్ క్యాంప్ ప్లానింగ్ & నెట్‌వర్క్' : '⛺ Service Camp Planning & Village Network'}
            </h1>
            <p className="text-[11px] text-slate-500 font-medium">
              {isTe
                ? 'హబ్ 4731 & 4732 బ్రాంచ్‌లు, మండలాలు, గ్రామాల దూరం & కస్టమర్ ట్రాక్టర్ల లెక్కలు'
                : 'Dealership Hubs 4731 & 4732 • Branch network • Village distance & live tractor counts'}
            </p>
          </div>
        </div>

        {/* Compact Stat Badges in a Row */}
        <div className="flex flex-wrap items-center gap-1.5 md:gap-2">
          <div className="px-2.5 py-1 bg-blue-50 border border-blue-200 rounded-lg flex items-center gap-1.5 text-blue-950 font-bold text-xs">
            <span className="text-[10px] text-blue-700 uppercase font-black">4731 Hub:</span>
            <span className="font-mono font-black text-blue-900">{stats.totalTractors4731}</span>
            <span className="text-[10px] text-slate-500 font-normal">🚜</span>
          </div>

          <div className="px-2.5 py-1 bg-teal-50 border border-teal-200 rounded-lg flex items-center gap-1.5 text-teal-950 font-bold text-xs">
            <span className="text-[10px] text-teal-700 uppercase font-black">4732 Hub:</span>
            <span className="font-mono font-black text-teal-900">{stats.totalTractors4732}</span>
            <span className="text-[10px] text-slate-500 font-normal">🚜</span>
          </div>

          <div className="px-2.5 py-1 bg-amber-50 border border-amber-200 rounded-lg flex items-center gap-1.5 text-amber-950 font-bold text-xs">
            <span className="text-[10px] text-amber-800 uppercase font-black">{isTe ? 'రాబోయేవి:' : 'Upcoming:'}</span>
            <span className="font-mono font-black text-amber-900">{stats.upcomingCamps}</span>
            <span className="text-[10px] text-slate-500 font-normal">⛺</span>
          </div>

          <div className="px-2.5 py-1 bg-emerald-50 border border-emerald-200 rounded-lg flex items-center gap-1.5 text-emerald-950 font-bold text-xs">
            <span className="text-[10px] text-emerald-800 uppercase font-black">{isTe ? 'పూర్తయినవి:' : 'Done:'}</span>
            <span className="font-mono font-black text-emerald-900">{stats.completedCamps}</span>
            <span className="text-[10px] text-slate-500 font-normal">✅</span>
          </div>

          {/* Quick Action Plan & Map Buttons */}
          <button
            type="button"
            onClick={() => setActiveTab('district_map')}
            className={`px-3 py-1.5 font-black text-xs rounded-lg shadow-2xs transition flex items-center gap-1 cursor-pointer ${
              activeTab === 'district_map'
                ? 'bg-amber-400 text-slate-950 ring-2 ring-amber-500'
                : 'bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300'
            }`}
            title="Open Krishna & NTR District Route Map & Distance Calculator"
          >
            <MapPin className="w-3.5 h-3.5 text-blue-900" />
            <span>{isTe ? '🗺️ జిల్లా రూట్ మ్యాప్' : '🗺️ Route Map'}</span>
          </button>

          <button
            type="button"
            onClick={() => handleOpenNewCampModal()}
            className="px-3 py-1.5 bg-blue-900 hover:bg-blue-800 text-white font-black text-xs rounded-lg shadow-2xs transition flex items-center gap-1 cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>{isTe ? 'కొత్త క్యాంప్' : 'Plan Camp'}</span>
          </button>
        </div>
      </div>

      {/* ========================================================================= */}
      {/* 2. THE COMPACT HORIZONTAL SELECTOR / FILTER BAR ("ADDA BOX") */}
      {/* Single Horizontal Strip with cascading dropdowns: Hub ▾ Branch ▾ Mandal ▾ Search */}
      {/* ========================================================================= */}
      <div className="bg-slate-900 text-white rounded-xl p-2.5 md:p-3 shadow-sm border border-slate-800">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-2.5">
          {/* Left: Cascading Dropdowns in a single horizontal row */}
          <div className="flex flex-wrap items-center gap-2 flex-1">
            {/* Primary View Switcher: Planning vs Scheduled vs District Route Map */}
            <div className="inline-flex rounded-lg bg-slate-800 p-0.5 border border-slate-700">
              <button
                type="button"
                onClick={() => setActiveTab('planning')}
                className={`px-3 py-1 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                  activeTab === 'planning' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-300 hover:text-white'
                }`}
              >
                <GitFork className="w-3.5 h-3.5" />
                <span>{isTe ? 'గ్రామాలు & ట్రీ' : 'Villages & Tree'}</span>
              </button>
              <button
                type="button"
                onClick={() => setActiveTab('camps_list')}
                className={`px-3 py-1 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                  activeTab === 'camps_list' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-300 hover:text-white'
                }`}
              >
                <Calendar className="w-3.5 h-3.5" />
                <span>{isTe ? 'ప్లాన్ చేసిన క్యాంప్స్' : 'Scheduled Camps'}</span>
                <span className="px-1.5 py-0.2 bg-amber-400 text-slate-950 rounded-full text-[10px] font-black">
                  {serviceCamps.length}
                </span>
              </button>
              <button
                type="button"
                onClick={() => setActiveTab('district_map')}
                className={`px-3 py-1 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                  activeTab === 'district_map' ? 'bg-amber-400 text-slate-950 font-black shadow-xs' : 'text-slate-300 hover:text-white'
                }`}
              >
                <MapPin className="w-3.5 h-3.5 text-rose-500" />
                <span>{isTe ? 'రూట్ & దూరాల మ్యాప్' : 'Route & Distance Map'}</span>
              </button>
            </div>

            <div className="h-5 w-px bg-slate-700 hidden sm:block" />

            {/* Dropdown 1: Dealership Hub */}
            <div className="flex items-center gap-1">
              <span className="text-[11px] font-bold text-slate-400 hidden xl:inline">{isTe ? 'హబ్:' : 'Hub:'}</span>
              <select
                value={selectedHub}
                onChange={e => handleHubChange(e.target.value as any)}
                className="bg-slate-800 border border-slate-700 hover:border-slate-600 text-white text-xs font-bold rounded-lg px-2.5 py-1.5 outline-none cursor-pointer"
              >
                <option value="all">{isTe ? '🏢 అన్నీ (All Hubs)' : '🏢 All Hubs (4731 & 4732)'}</option>
                <option value="4731">4731 Hub (Tiruvuru, Nuzvidu, Nandigama)</option>
                <option value="4732">4732 Hub (Gudivada, Poranki, Machilipatnam)</option>
              </select>
            </div>

            {/* Dropdown 2: Branch */}
            <div className="flex items-center gap-1">
              <span className="text-[11px] font-bold text-slate-400 hidden xl:inline">{isTe ? 'బ్రాంచ్:' : 'Branch:'}</span>
              <select
                value={selectedBranch}
                onChange={e => handleBranchChange(e.target.value)}
                className="bg-slate-800 border border-slate-700 hover:border-slate-600 text-white text-xs font-bold rounded-lg px-2.5 py-1.5 outline-none cursor-pointer"
              >
                <option value="all">{isTe ? '🏬 అన్ని బ్రాంచ్‌లు (All Branches)' : '🏬 All Branches'}</option>
                {availableBranches.map(b => (
                  <option key={b.id} value={b.name}>
                    {b.name} ({b.teluguName}) • {b.dealershipCode}
                  </option>
                ))}
              </select>
            </div>

            {/* Dropdown 3: Mandal */}
            <div className="flex items-center gap-1">
              <span className="text-[11px] font-bold text-slate-400 hidden xl:inline">{isTe ? 'మండలం:' : 'Mandal:'}</span>
              <select
                value={selectedMandal}
                onChange={e => {
                  setSelectedMandal(e.target.value);
                  setTablePage(1);
                  if (e.target.value !== 'all') {
                    // Auto-expand this mandal in tree
                    const mKey = Object.keys(expandedMandals).find(k => k.endsWith(`-${e.target.value}`));
                    if (mKey) setExpandedMandals(prev => ({ ...prev, [mKey]: true }));
                  }
                }}
                className="bg-slate-800 border border-slate-700 hover:border-slate-600 text-white text-xs font-bold rounded-lg px-2.5 py-1.5 outline-none cursor-pointer max-w-[170px]"
              >
                <option value="all">{isTe ? '📁 అన్ని మండలాలు (All Mandals)' : '📁 All Mandals'}</option>
                {availableMandals.map(m => (
                  <option key={m} value={m}>
                    {m} Mandal
                  </option>
                ))}
              </select>
            </div>

            {/* Quick Search */}
            <div className="relative flex-1 min-w-[160px] max-w-xs">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchQuery}
                onChange={e => {
                  setSearchQuery(e.target.value);
                  setTablePage(1);
                }}
                placeholder={isTe ? 'గ్రామం / మండలం వెతకండి...' : 'Search village / mandal...'}
                className="w-full bg-slate-800 border border-slate-700 text-white text-xs rounded-lg pl-8 pr-7 py-1.5 outline-none focus:border-blue-500 placeholder:text-slate-400 font-medium"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </div>
          </div>

          {/* Right: View style switcher (Tree vs Table) & Tree Controls */}
          {activeTab === 'planning' && (
            <div className="flex items-center gap-2 shrink-0 self-end lg:self-center">
              {/* View Style Switcher (Tree 🌳 vs Table 📋) */}
              <div className="inline-flex rounded-lg bg-slate-800 p-0.5 border border-slate-700">
                <button
                  type="button"
                  onClick={() => setViewStyle('tree')}
                  className={`px-2.5 py-1 rounded text-xs font-bold transition flex items-center gap-1 cursor-pointer ${
                    viewStyle === 'tree' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-400 hover:text-white'
                  }`}
                  title={isTe ? 'ట్రీ వ్యూ (Tree View)' : 'Tree View'}
                >
                  <GitFork className="w-3.5 h-3.5" />
                  <span className="hidden sm:inline">{isTe ? 'ట్రీ వ్యూ' : 'Tree'}</span>
                </button>
                <button
                  type="button"
                  onClick={() => setViewStyle('table')}
                  className={`px-2.5 py-1 rounded text-xs font-bold transition flex items-center gap-1 cursor-pointer ${
                    viewStyle === 'table' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-400 hover:text-white'
                  }`}
                  title={isTe ? 'టేబుల్ వ్యూ (Table View)' : 'Table View'}
                >
                  <List className="w-3.5 h-3.5" />
                  <span className="hidden sm:inline">{isTe ? 'టేబుల్ వ్యూ' : 'Table'}</span>
                  <span className="text-[10px] opacity-80 font-mono">({flatVillagesList.length})</span>
                </button>
              </div>

              {viewStyle === 'tree' && (
                <div className="flex items-center gap-1">
                  <button
                    type="button"
                    onClick={expandAll}
                    className="px-2 py-1 text-[11px] font-bold text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 transition cursor-pointer"
                    title={isTe ? 'అన్నీ తెరవండి' : 'Expand All'}
                  >
                    + {isTe ? 'అన్నీ' : 'All'}
                  </button>
                  <button
                    type="button"
                    onClick={collapseAll}
                    className="px-2 py-1 text-[11px] font-bold text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 transition cursor-pointer"
                    title={isTe ? 'అన్నీ మూసివేయండి' : 'Collapse All'}
                  >
                    - {isTe ? 'మూయి' : 'Close'}
                  </button>
                </div>
              )}

              <button
                type="button"
                onClick={exportServiceCampsExcel}
                className="p-1.5 text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg border border-slate-700 transition cursor-pointer"
                title="Export Scheduled Camps to CSV"
              >
                <Download className="w-3.5 h-3.5 text-emerald-400" />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* ========================================================================= */}
      {/* 3. MAIN CONTENT: 2-COLUMN SPLIT (TREE ON LEFT, PEDDA BOX FOR CUSTOMER DATA ON RIGHT) */}
      {/* ========================================================================= */}
      {activeTab === 'planning' && viewStyle === 'tree' && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 items-start">
          {/* LEFT COLUMN: MANDALS & VILLAGES TREE DIRECTORY (COMPACT OKA PAKKANA) */}
          <div className="lg:col-span-5 xl:col-span-5 flex flex-col space-y-2">
            <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden flex flex-col">
              <div className="p-2.5 bg-slate-900 text-white flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <GitFork className="w-4 h-4 text-amber-400" />
                  <span className="font-black text-xs">
                    {isTe ? '🌳 మండలాలు & గ్రామాల ట్రీ' : '🌳 Mandals & Villages Tree'}
                  </span>
                </div>
                <span className="text-[10px] text-slate-300">
                  {filteredDealerships.reduce((acc, h) => acc + h.branches.reduce((bAcc, b) => bAcc + b.mandals.reduce((mAcc, m) => mAcc + m.villages.length, 0), 0), 0)} {isTe ? 'గ్రామాలు' : 'Villages'}
                </span>
              </div>

              <div className="h-[calc(100vh-230px)] min-h-[580px] max-h-[780px] overflow-y-auto divide-y divide-slate-200">
          {filteredDealerships.length === 0 ? (
            <div className="p-8 text-center text-slate-500 space-y-2">
              <Search className="w-8 h-8 text-slate-400 mx-auto" />
              <p className="text-sm font-bold text-slate-700">
                {isTe ? 'ఎటువంటి ఫలితాలు కనుగొనబడలేదు' : 'No branches, mandals or villages match your filter.'}
              </p>
              <button
                type="button"
                onClick={() => {
                  setSelectedHub('all');
                  setSelectedBranch('all');
                  setSelectedMandal('all');
                  setSearchQuery('');
                }}
                className="px-3 py-1 bg-blue-900 text-white rounded text-xs font-bold"
              >
                {isTe ? 'ఫిల్టర్లు రీసెట్ చేయండి' : 'Reset Filters'}
              </button>
            </div>
          ) : (
            <div className="divide-y divide-slate-200">
              {filteredDealerships.map(hub => {
                const isHubExpanded = expandedHubs[hub.code] ?? true;
                const is4731 = hub.code === '4731';

                // Count total tractors in this hub
                let hubTractors = 0;
                let hubVillagesCount = 0;
                hub.branches.forEach(b => {
                  b.mandals.forEach(m => {
                    hubVillagesCount += m.villages.length;
                    m.villages.forEach(v => {
                      hubTractors += getTractorCountForVillage(v.name, m.name);
                    });
                  });
                });

                return (
                  <div key={hub.code} className="bg-white">
                    {/* LEVEL 1: HUB COMPACT ROW */}
                    <div
                      onClick={() => toggleHub(hub.code)}
                      className={`px-3 py-2 flex items-center justify-between cursor-pointer select-none transition border-b border-slate-200 ${
                        is4731 ? 'bg-blue-950 text-white' : 'bg-slate-900 text-white'
                      }`}
                    >
                      <div className="flex items-center gap-2">
                        <span className="p-1 rounded bg-white/10 hover:bg-white/20">
                          {isHubExpanded ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />}
                        </span>
                        <span className="px-1.5 py-0.2 bg-amber-400 text-slate-950 font-mono font-black text-[10px] rounded uppercase">
                          CODE {hub.code}
                        </span>
                        <span className="font-black text-xs md:text-sm">
                          {isTe ? hub.teluguName : hub.name}
                        </span>
                      </div>

                      <div className="flex items-center gap-2">
                        <span className="text-[11px] text-slate-300 font-semibold hidden sm:inline">
                          {hub.branches.length} {isTe ? 'బ్రాంచ్‌లు' : 'Branches'} • {hubVillagesCount} {isTe ? 'గ్రామాలు' : 'Villages'}
                        </span>
                        <span className="px-2 py-0.5 bg-white/20 rounded text-[11px] font-mono font-black text-amber-300">
                          🚜 {hubTractors}
                        </span>
                        <button
                          type="button"
                          onClick={e => {
                            e.stopPropagation();
                            handleOpenNewCampModal({ dealershipCode: hub.code });
                          }}
                          className="px-2 py-0.5 bg-amber-400 hover:bg-amber-300 text-slate-950 text-[11px] font-black rounded transition"
                        >
                          + {isTe ? 'క్యాంప్' : 'Plan'}
                        </button>
                      </div>
                    </div>

                    {/* LEVEL 2: BRANCHES */}
                    {isHubExpanded && (
                      <div className="divide-y divide-slate-100 bg-slate-50/50">
                        {hub.branches.map(branch => {
                          const isBranchExpanded = expandedBranches[branch.id] ?? false;

                          let branchTractors = 0;
                          let branchVillagesCount = 0;
                          branch.mandals.forEach(m => {
                            branchVillagesCount += m.villages.length;
                            m.villages.forEach(v => {
                              branchTractors += getTractorCountForVillage(v.name, m.name);
                            });
                          });

                          return (
                            <div key={branch.id}>
                              {/* Branch Slim Row */}
                              <div
                                onClick={() => toggleBranch(branch.id)}
                                className={`px-4 py-2 flex items-center justify-between cursor-pointer select-none transition ${
                                  isBranchExpanded ? 'bg-blue-50/80 border-l-4 border-blue-900' : 'bg-white hover:bg-slate-100/70 border-l-4 border-transparent'
                                }`}
                              >
                                <div className="flex items-center gap-2.5">
                                  <span className="text-slate-500">
                                    {isBranchExpanded ? (
                                      <FolderOpen className="w-4 h-4 text-blue-900" />
                                    ) : (
                                      <Folder className="w-4 h-4 text-slate-500" />
                                    )}
                                  </span>
                                  <div>
                                    <span className="font-black text-xs text-slate-900">
                                      {branch.name} ({branch.teluguName})
                                    </span>
                                    <span className="text-[10px] text-slate-500 font-medium ml-2 hidden md:inline">
                                      📍 {branch.hubAddress}
                                    </span>
                                  </div>
                                </div>

                                <div className="flex items-center gap-2">
                                  <span className="text-[10px] text-slate-500 font-semibold">
                                    {branch.mandals.length} {isTe ? 'మండలాలు' : 'Mandals'} ({branchVillagesCount} {isTe ? 'గ్రామాలు' : 'Villages'})
                                  </span>
                                  <span className="px-2 py-0.5 bg-blue-100 text-blue-900 rounded text-[10px] font-mono font-black">
                                    🚜 {branchTractors}
                                  </span>
                                  <button
                                    type="button"
                                    onClick={e => {
                                      e.stopPropagation();
                                      handleOpenNewCampModal({
                                        dealershipCode: hub.code,
                                        branch: branch.name
                                      });
                                    }}
                                    className="px-2 py-0.5 bg-blue-900 hover:bg-blue-800 text-white text-[10px] font-bold rounded transition"
                                  >
                                    + {isTe ? 'క్యాంప్' : 'Plan'}
                                  </button>
                                </div>
                              </div>

                              {/* LEVEL 3: MANDALS */}
                              {isBranchExpanded && (
                                <div className="pl-6 pr-2 py-1 space-y-1 bg-slate-100/60 border-t border-slate-100">
                                  {branch.mandals.map(mandal => {
                                    const mandalKey = `${branch.id}-${mandal.name}`;
                                    const isMandalExpanded = expandedMandals[mandalKey] ?? false;

                                    let mandalTractors = 0;
                                    mandal.villages.forEach(v => {
                                      mandalTractors += getTractorCountForVillage(v.name, mandal.name);
                                    });

                                    return (
                                      <div key={mandal.name} className="bg-white rounded-lg border border-slate-200 overflow-hidden shadow-2xs">
                                        {/* Mandal Header Bar */}
                                        <div
                                          onClick={() => toggleMandal(mandalKey)}
                                          className={`px-3 py-1.5 flex items-center justify-between cursor-pointer select-none transition ${
                                            isMandalExpanded ? 'bg-slate-100 border-b border-slate-200' : 'bg-white hover:bg-slate-50'
                                          }`}
                                        >
                                          <div className="flex items-center gap-2">
                                            <span className="p-0.5 text-slate-600">
                                              {isMandalExpanded ? (
                                                <ChevronDown className="w-3.5 h-3.5" />
                                              ) : (
                                                <ChevronRight className="w-3.5 h-3.5" />
                                              )}
                                            </span>
                                            <span className="font-black text-xs text-slate-800">
                                              {mandal.name} Mandal {mandal.teluguName ? `(${mandal.teluguName})` : ''}
                                            </span>
                                            {mandal.distanceFromBranchKm > 0 && (
                                              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-1.5 py-0.2 rounded border border-slate-200">
                                                📍 {mandal.distanceFromBranchKm} km from {branch.name}
                                              </span>
                                            )}
                                          </div>

                                          <div className="flex items-center gap-1.5">
                                            <span className="text-[10px] text-slate-500">
                                              {mandal.villages.length} {isTe ? 'గ్రామాలు' : 'Villages'}
                                            </span>
                                            <span className="px-1.5 py-0.2 bg-emerald-50 text-emerald-900 border border-emerald-200 rounded text-[10px] font-mono font-black">
                                              🚜 {mandalTractors}
                                            </span>
                                            <button
                                              type="button"
                                              onClick={e => {
                                                e.stopPropagation();
                                                handleOpenNewCampModal({
                                                  dealershipCode: hub.code,
                                                  branch: branch.name,
                                                  mandal: mandal.name
                                                });
                                              }}
                                              className="px-1.5 py-0.5 bg-blue-50 text-blue-900 hover:bg-blue-100 border border-blue-200 text-[10px] font-bold rounded"
                                            >
                                              + {isTe ? 'ప్లాన్' : 'Plan'}
                                            </button>
                                          </div>
                                        </div>

                                        {/* LEVEL 4: COMPACT VILLAGE LINE ITEMS (Directory Style with Inline Customer Expansion) */}
                                        {isMandalExpanded && (
                                          <div className="divide-y divide-slate-100 bg-white">
                                            {mandal.villages.map((village, vIdx) => {
                                              const villageKey = `${branch.id}-${mandal.name}-${village.name}`;
                                              const isSelected = selectedVillageKey === villageKey || (
                                                activeVillageDetails?.villageName === village.name &&
                                                activeVillageDetails?.mandalName === mandal.name
                                              );
                                              const villageCusts = getCustomersForVillage(village.name, mandal.name);
                                              const tractorCount = villageCusts.length || getTractorCountForVillage(village.name, mandal.name);
                                              const directionsUrl = getGoogleMapsDirectionsUrl(branch.name, village.name, mandal.name);

                                              return (
                                                <div
                                                  key={village.name}
                                                  onClick={() => handleSelectVillage(village, mandal, branch, hub.code)}
                                                  className={`px-3 py-2 flex items-center justify-between gap-2 transition text-xs cursor-pointer select-none border-b border-slate-100 ${
                                                    isSelected
                                                      ? 'bg-blue-900 text-white font-bold shadow-xs'
                                                      : 'hover:bg-blue-50/50 bg-white text-slate-800'
                                                  }`}
                                                  title={isTe ? 'కస్టమర్ డేటాని పక్కన ఉన్న పెద్ద బాక్స్‌లో చూడటానికి క్లిక్ చేయండి' : 'Click to view customer details in the right side box'}
                                                >
                                                  {/* Left: Village S.No & Name */}
                                                  <div className="flex items-center gap-1.5 min-w-0">
                                                    <span className={`text-[10px] font-mono w-4 shrink-0 text-right ${isSelected ? 'text-blue-200' : 'text-slate-400'}`}>
                                                      {vIdx + 1}.
                                                    </span>
                                                    <div className="truncate">
                                                      <div className="font-bold truncate leading-tight">
                                                        {village.name}
                                                      </div>
                                                      {village.teluguName && (
                                                        <div className={`text-[10px] truncate ${isSelected ? 'text-blue-200' : 'text-slate-500'}`}>
                                                          {village.teluguName}
                                                        </div>
                                                      )}
                                                    </div>
                                                  </div>

                                                  {/* Middle & Right: Distance & Live Tractor Count */}
                                                  <div className="flex items-center gap-1.5 shrink-0">
                                                    {/* Distance Badge */}
                                                    <span className={`text-[10px] font-medium px-1.5 py-0.5 rounded ${
                                                      isSelected ? 'bg-blue-800 text-blue-100' : 'bg-slate-100 text-slate-600'
                                                    }`}>
                                                      📍 {village.distanceKm} km
                                                    </span>

                                                    {/* Live Tractor Count */}
                                                    <span className={`px-2 py-0.5 rounded-full text-[10px] font-mono font-black flex items-center gap-1 ${
                                                      tractorCount > 0
                                                        ? isSelected
                                                          ? 'bg-amber-400 text-slate-950 ring-1 ring-white'
                                                          : 'bg-blue-100 text-blue-900'
                                                        : isSelected
                                                          ? 'bg-blue-800 text-blue-200'
                                                          : 'bg-slate-100 text-slate-400'
                                                    }`}>
                                                      <span>🚜</span>
                                                      <span>{tractorCount}</span>
                                                    </span>

                                                    <ChevronRight className={`w-3.5 h-3.5 ${isSelected ? 'text-white' : 'text-slate-400'}`} />
                                                  </div>
                                                </div>
                                              );
                                            })}
                                          </div>
                                        )}
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
              </div>
            </div>
          </div>

          {/* =================================================================== */}
          {/* RIGHT COLUMN: THE "PEDDA BOX" (LARGE DEDICATED CUSTOMER DATA CONTAINER) */}
          {/* =================================================================== */}
          <div className="lg:col-span-7 xl:col-span-7 sticky top-3">
            <div className="bg-white rounded-xl border-2 border-slate-300 shadow-md flex flex-col h-[calc(100vh-210px)] min-h-[580px] max-h-[820px] overflow-hidden">
              {activeVillageDetails ? (
                <div className="flex-1 flex flex-col h-full overflow-hidden">
                  {/* Top Bar of Pedda Box */}
                  <div className="bg-slate-900 text-white px-3.5 py-2.5 border-b border-slate-800 shrink-0 flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <span className="p-1.5 bg-blue-600 text-white rounded-lg">
                        <Users className="w-4 h-4" />
                      </span>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <h2 className="text-sm md:text-base font-black text-white">
                            {activeVillageDetails.villageName} {activeVillageDetails.villageTelugu ? `(${activeVillageDetails.villageTelugu})` : ''}
                          </h2>
                          <span className="px-1.5 py-0.5 bg-blue-800 text-blue-100 rounded text-[10px] font-bold">
                            {activeVillageDetails.mandalName} Mandal
                          </span>
                          <span className="px-1.5 py-0.5 bg-slate-800 text-amber-300 rounded text-[10px] font-bold">
                            {activeVillageDetails.branchName} ({activeVillageDetails.hubCode} Hub)
                          </span>
                        </div>
                        <div className="text-[11px] text-slate-400 font-medium flex items-center gap-2 mt-0.5">
                          <span>📍 {activeVillageDetails.distanceKm} km ({activeVillageDetails.approxTravelTime || '20 min'})</span>
                          <span>•</span>
                          <span className="font-bold text-amber-400 font-mono">
                            🚜 {getCustomersForVillage(activeVillageDetails.villageName, activeVillageDetails.mandalName).length} {isTe ? 'కస్టమర్ ట్రాక్టర్లు' : 'Customers'}
                          </span>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-1.5">
                      <a
                        href={getGoogleMapsDirectionsUrl(activeVillageDetails.branchName, activeVillageDetails.villageName, activeVillageDetails.mandalName)}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="px-2 py-1 text-[11px] font-bold text-blue-300 bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 transition flex items-center gap-1 cursor-pointer"
                        title="Google Maps"
                      >
                        <MapIcon className="w-3 h-3" />
                        <span className="hidden sm:inline">{isTe ? 'రూట్' : 'Maps'}</span>
                      </a>
                      <button
                        type="button"
                        onClick={() => handleOpenNewCampModal({
                          dealershipCode: activeVillageDetails.hubCode as any,
                          branch: activeVillageDetails.branchName,
                          mandal: activeVillageDetails.mandalName,
                          village: activeVillageDetails.villageName
                        })}
                        className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded text-[11px] transition shadow-xs flex items-center gap-1 cursor-pointer"
                      >
                        <Plus className="w-3 h-3" />
                        <span>{isTe ? 'క్యాంప్ ప్లాన్' : 'Plan Camp'}</span>
                      </button>
                    </div>
                  </div>

                  {/* Body: Dedicated Scrollable Customer Data Container (The user scrolls ONLY inside this box) */}
                  <div className="flex-1 overflow-y-auto bg-slate-50/50">
                    <VillageCustomerInline
                      villageName={activeVillageDetails.villageName}
                      villageTelugu={activeVillageDetails.villageTelugu}
                      mandalName={activeVillageDetails.mandalName}
                      mandalTelugu={activeVillageDetails.mandalTelugu}
                      branchName={activeVillageDetails.branchName}
                      hubCode={activeVillageDetails.hubCode}
                      distanceKm={activeVillageDetails.distanceKm}
                      approxTravelTime={activeVillageDetails.approxTravelTime}
                      customers={getCustomersForVillage(activeVillageDetails.villageName, activeVillageDetails.mandalName)}
                      isTe={isTe}
                      onPlanCamp={initialData => handleOpenNewCampModal(initialData)}
                      onStartBroadcast={data => setBroadcastData(data)}
                      contactedCustomers={contactedCustomers}
                      onMarkContacted={handleMarkContacted}
                      showToast={showToast}
                    />
                  </div>
                </div>
              ) : (
                <div className="flex-1 flex flex-col items-center justify-center p-8 text-center text-slate-500 space-y-3">
                  <Users className="w-12 h-12 text-slate-300" />
                  <div className="max-w-md">
                    <h3 className="font-bold text-slate-800 text-sm">
                      {isTe ? '👈 ఎడమవైపు ట్రీ నుండి ఏదైనా గ్రామాన్ని ఎంచుకోండి' : 'Select a village from the tree on the left'}
                    </h3>
                    <p className="text-xs text-slate-500 mt-1">
                      {isTe
                        ? 'ఆ గ్రామంలోని కస్టమర్ల వివరాలు, మోడల్స్, ఫోన్ నంబర్లు మరియు వాట్సాప్ బ్రాడ్‌కాస్ట్ ఈ పెద్ద బాక్స్‌లో కనిపిస్తాయి. మీరు పేజీ మొత్తం స్క్రోల్ చేయకుండా ఇక్కడే స్క్రోల్ చేసి చూసుకోవచ్చు.'
                        : 'Customer details, models, phone numbers and WhatsApp features will appear in this dedicated box with independent scrolling.'}
                    </p>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 4. COMPACT TABLE VIEW (High-density list of all filtered villages) */}
      {/* ========================================================================= */}
      {activeTab === 'planning' && viewStyle === 'table' && (
        <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden space-y-2">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200 uppercase text-[10px]">
                  <th className="p-2.5 text-center">#</th>
                  <th className="p-2.5">{isTe ? 'గ్రామం (Village)' : 'Village Name'}</th>
                  <th className="p-2.5">{isTe ? 'మండలం (Mandal)' : 'Mandal'}</th>
                  <th className="p-2.5">{isTe ? 'బ్రాంచ్ & హబ్' : 'Branch & Hub'}</th>
                  <th className="p-2.5">{isTe ? 'బ్రాంచ్ నుండి దూరం' : 'Distance from Branch'}</th>
                  <th className="p-2.5 text-center">{isTe ? 'కస్టమర్ ట్రాక్టర్లు' : 'Tractor Count'}</th>
                  <th className="p-2.5 text-center">{isTe ? 'గూగుల్ మ్యాప్స్' : 'Maps'}</th>
                  <th className="p-2.5 text-center">{isTe ? 'చర్య' : 'Action'}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 bg-white">
                {flatVillagesList.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="p-8 text-center text-slate-500 font-bold">
                      {isTe ? 'ఎటువంటి గ్రామాలు కనుగొనబడలేదు.' : 'No villages match the selected filter.'}
                    </td>
                  </tr>
                ) : (
                  paginatedVillages.map((item, idx) => {
                    const globalIdx = (tablePage - 1) * pageSize + idx + 1;
                    const villageKey = `table-${item.branchName}-${item.mandalName}-${item.villageName}`;
                    const isVillageExpanded = expandedVillages[villageKey] ?? false;
                    const villageCusts = getCustomersForVillage(item.villageName, item.mandalName);

                    return (
                      <React.Fragment key={`${item.branchName}-${item.mandalName}-${item.villageName}`}>
                        <tr
                          onClick={() => toggleVillage(villageKey)}
                          className={`hover:bg-blue-50/40 transition cursor-pointer select-none ${
                            isVillageExpanded ? 'bg-blue-50/70 border-l-4 border-blue-900 font-bold' : ''
                          }`}
                        >
                          <td className="p-2.5 text-center font-mono text-slate-400 text-[11px]">{globalIdx}</td>
                          <td className="p-2.5">
                            <div className="flex items-center gap-1.5">
                              <span className="text-slate-400">
                                {isVillageExpanded ? (
                                  <ChevronDown className="w-3.5 h-3.5 text-blue-900" />
                                ) : (
                                  <ChevronRight className="w-3.5 h-3.5 text-slate-400" />
                                )}
                              </span>
                              <div>
                                <div className="font-bold text-slate-900">{item.villageName}</div>
                                {item.villageTelugu && (
                                  <div className="text-[10px] text-slate-500 font-medium">{item.villageTelugu}</div>
                                )}
                              </div>
                            </div>
                          </td>
                          <td className="p-2.5 font-semibold text-slate-800">
                            {item.mandalName} {item.mandalTelugu ? `(${item.mandalTelugu})` : ''}
                          </td>
                          <td className="p-2.5">
                            <span className="font-bold text-slate-800">{item.branchName}</span>
                            <span className="ml-1 text-[10px] font-mono px-1 py-0.2 rounded bg-slate-100 text-slate-600 font-bold">
                              {item.hubCode}
                            </span>
                          </td>
                          <td className="p-2.5 text-slate-600 font-medium">
                            📍 {item.distanceKm} km ({item.approxTravelTime || '20 min'})
                          </td>
                          <td className="p-2.5 text-center">
                            <button
                              type="button"
                              onClick={e => {
                                e.stopPropagation();
                                toggleVillage(villageKey);
                              }}
                              className={`px-2.5 py-1 rounded text-xs font-mono font-bold transition inline-flex items-center gap-1 cursor-pointer ${
                                item.tractorCount > 0
                                  ? isVillageExpanded
                                    ? 'bg-blue-950 text-white ring-2 ring-blue-400 shadow-2xs'
                                    : 'bg-blue-900 text-white hover:bg-blue-800 shadow-2xs'
                                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                              }`}
                            >
                              <span>🚜 {item.tractorCount}</span>
                              <span className="text-[9px] uppercase font-bold">
                                {isVillageExpanded ? (isTe ? 'మూయి' : 'Hide') : (isTe ? 'చూడు' : 'View')}
                              </span>
                            </button>
                          </td>
                          <td className="p-2.5 text-center">
                            <a
                              href={item.directionsUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              onClick={e => e.stopPropagation()}
                              className="p-1.5 inline-flex text-blue-900 hover:text-blue-950 bg-blue-50 hover:bg-blue-100 rounded transition"
                              title={`Google Maps directions from ${item.branchName}`}
                            >
                              <MapIcon className="w-3.5 h-3.5" />
                            </a>
                          </td>
                          <td className="p-2.5 text-center">
                            <button
                              type="button"
                              onClick={e => {
                                e.stopPropagation();
                                handleOpenNewCampModal({
                                  dealershipCode: item.hubCode,
                                  branch: item.branchName,
                                  mandal: item.mandalName,
                                  village: item.villageName
                                });
                              }}
                              className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded text-xs transition cursor-pointer"
                            >
                              ⛺ {isTe ? 'ప్లాన్' : 'Plan'}
                            </button>
                          </td>
                        </tr>

                        {/* Inline Village Customer Intelligence View inside Table */}
                        {isVillageExpanded && (
                          <tr>
                            <td colSpan={8} className="p-0 border-b-2 border-blue-900">
                              <VillageCustomerInline
                                villageName={item.villageName}
                                villageTelugu={item.villageTelugu}
                                mandalName={item.mandalName}
                                mandalTelugu={item.mandalTelugu}
                                branchName={item.branchName}
                                hubCode={item.hubCode}
                                distanceKm={item.distanceKm}
                                approxTravelTime={item.approxTravelTime}
                                customers={villageCusts}
                                isTe={isTe}
                                onPlanCamp={initialData => handleOpenNewCampModal(initialData)}
                                onStartBroadcast={data => setBroadcastData(data)}
                                contactedCustomers={contactedCustomers}
                                onMarkContacted={handleMarkContacted}
                                onClose={() => toggleVillage(villageKey)}
                                showToast={showToast}
                              />
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Table Pagination Bar */}
          {flatVillagesList.length > pageSize && (
            <div className="p-2.5 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs font-semibold text-slate-700">
              <div className="text-[11px] text-slate-500">
                Showing {(tablePage - 1) * pageSize + 1} to {Math.min(tablePage * pageSize, flatVillagesList.length)} of {flatVillagesList.length} villages
              </div>

              <div className="flex items-center gap-1">
                <button
                  type="button"
                  disabled={tablePage <= 1}
                  onClick={() => setTablePage(p => Math.max(1, p - 1))}
                  className="px-2.5 py-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  ‹ {isTe ? 'మునుపటి' : 'Prev'}
                </button>
                <span className="px-2.5 py-1 bg-white border border-slate-300 rounded font-bold font-mono">
                  {tablePage} / {totalTablePages}
                </span>
                <button
                  type="button"
                  disabled={tablePage >= totalTablePages}
                  onClick={() => setTablePage(p => Math.min(totalTablePages, p + 1))}
                  className="px-2.5 py-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  {isTe ? 'తరువాతి' : 'Next'} ›
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* 5. TAB 2: SCHEDULED SERVICE CAMPS LIST */}
      {/* ========================================================================= */}
      {activeTab === 'camps_list' && (
        <div className="bg-white rounded-xl border border-slate-200 shadow-2xs p-3 space-y-3">
          {/* Status Filter Tabs */}
          <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-slate-200">
            <div className="flex flex-wrap items-center gap-1.5">
              {(['all', 'Upcoming', 'In Progress', 'Completed', 'Cancelled'] as const).map(st => (
                <button
                  key={st}
                  type="button"
                  onClick={() => setCampStatusFilter(st)}
                  className={`px-3 py-1 text-xs font-bold rounded-lg transition cursor-pointer ${
                    campStatusFilter === st
                      ? 'bg-blue-900 text-white shadow-2xs'
                      : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  {st === 'all'
                    ? isTe ? 'అన్నీ (All)' : 'All'
                    : st === 'Upcoming'
                    ? isTe ? 'రాబోయేవి (Upcoming)' : 'Upcoming'
                    : st === 'In Progress'
                    ? isTe ? 'జరుగుతున్నవి' : 'In Progress'
                    : st === 'Completed'
                    ? isTe ? 'పూర్తయినవి' : 'Completed'
                    : isTe ? 'రద్దు' : 'Cancelled'}
                  {' '}
                  <span className="ml-1 text-[10px] opacity-80">
                    ({st === 'all' ? serviceCamps.length : serviceCamps.filter(c => c.status === st).length})
                  </span>
                </button>
              ))}
            </div>

            <button
              type="button"
              onClick={() => handleOpenNewCampModal()}
              className="px-3 py-1 bg-blue-900 text-white text-xs font-bold rounded-lg hover:bg-blue-800 transition flex items-center gap-1 cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>{isTe ? 'కొత్త క్యాంప్' : 'Add Camp'}</span>
            </button>
          </div>

          {/* Compact Camps Table */}
          {filteredCamps.length === 0 ? (
            <div className="p-8 text-center space-y-2">
              <Calendar className="w-10 h-10 text-slate-300 mx-auto" />
              <p className="text-sm font-bold text-slate-700">
                {isTe ? 'ఎటువంటి సర్వీస్ క్యాంప్స్ కనుగొనబడలేదు' : 'No service camps found'}
              </p>
              <button
                type="button"
                onClick={() => handleOpenNewCampModal()}
                className="px-3 py-1.5 bg-blue-900 text-white text-xs font-bold rounded-lg"
              >
                + {isTe ? 'ఇప్పుడే ప్లాన్ చేయండి' : 'Plan Camp Now'}
              </button>
            </div>
          ) : (
            <div className="overflow-x-auto rounded-lg border border-slate-200">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200 uppercase text-[10px]">
                    <th className="p-2.5">#</th>
                    <th className="p-2.5">{isTe ? 'తేదీ & స్థితి' : 'Date & Status'}</th>
                    <th className="p-2.5">{isTe ? 'గ్రామం / మండలం / బ్రాంచ్' : 'Village / Location'}</th>
                    <th className="p-2.5">{isTe ? 'టార్గెట్ / హాజరైనవి' : 'Target / Attended'}</th>
                    <th className="p-2.5">{isTe ? 'సూపర్‌వైజర్ & మెకానిక్' : 'Team'}</th>
                    <th className="p-2.5">{isTe ? 'సమన్వయకర్త & ఆఫర్లు' : 'Contact & Offers'}</th>
                    <th className="p-2.5 text-center">{isTe ? 'చర్యలు' : 'Actions'}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 bg-white">
                  {filteredCamps.map((camp, idx) => {
                    const villageCusts = getCustomersForVillage(camp.village, camp.mandal);
                    const directionsUrl = getGoogleMapsDirectionsUrl(camp.branch, camp.village, camp.mandal);

                    return (
                      <tr key={camp.id} className="hover:bg-slate-50 transition">
                        <td className="p-2.5 font-mono text-slate-400">{idx + 1}</td>
                        <td className="p-2.5">
                          <div className="font-bold text-slate-900 flex items-center gap-1">
                            <CalendarDays className="w-3.5 h-3.5 text-blue-900 shrink-0" />
                            <span>{camp.campDate}</span>
                          </div>
                          <span
                            className={`inline-block mt-0.5 px-2 py-0.2 rounded-full text-[9px] font-black uppercase ${
                              camp.status === 'Upcoming'
                                ? 'bg-amber-100 text-amber-900 border border-amber-300'
                                : camp.status === 'In Progress'
                                ? 'bg-blue-100 text-blue-900 border border-blue-300 animate-pulse'
                                : camp.status === 'Completed'
                                ? 'bg-emerald-100 text-emerald-900 border border-emerald-300'
                                : 'bg-slate-200 text-slate-700'
                            }`}
                          >
                            {camp.status}
                          </span>
                        </td>

                        <td className="p-2.5">
                          <div className="font-bold text-slate-900">
                            {camp.village}, {camp.mandal}
                          </div>
                          <div className="text-[10px] text-slate-500 font-semibold">
                            {camp.dealershipCode} Hub • {camp.branch}
                          </div>
                          {camp.venue && (
                            <div className="text-[10px] text-slate-600 mt-0.5 flex items-center gap-1">
                              <MapPin className="w-3 h-3 text-red-500 shrink-0" />
                              <span>{camp.venue}</span>
                            </div>
                          )}
                        </td>

                        <td className="p-2.5">
                          <div className="font-bold text-slate-900">🎯 Target: {camp.targetTractors}</div>
                          <div className="text-[11px] font-bold text-emerald-700">
                            ✅ Attended: {camp.attendedCount || 0}
                          </div>
                        </td>

                        <td className="p-2.5 space-y-0.5">
                          <div className="text-slate-700">
                            <span className="font-bold text-slate-500 text-[10px]">Sup:</span> {camp.supervisor || 'N/A'}
                          </div>
                          <div className="text-slate-700">
                            <span className="font-bold text-slate-500 text-[10px]">Mech:</span> {camp.mechanic || 'N/A'}
                          </div>
                        </td>

                        <td className="p-2.5 max-w-xs">
                          {camp.contactPerson && (
                            <div className="font-bold text-slate-800">
                              {camp.contactPerson} {camp.contactPhone ? `(${camp.contactPhone})` : ''}
                            </div>
                          )}
                          {camp.offers && (
                            <div className="text-[10px] text-amber-900 line-clamp-1 mt-0.5">
                              🎁 {camp.offers}
                            </div>
                          )}
                        </td>

                        <td className="p-2.5 text-center">
                          <div className="flex items-center justify-center gap-1">
                            {/* Broadcast / Farmer outreach for this camp */}
                            <button
                              type="button"
                              onClick={() => {
                                const defaultMsg = isTe
                                  ? `నమస్కారం! శ్రీ గాయత్రి ఆటోమోటివ్స్ (ఈచర్ ట్రాక్టర్స్ డీలర్షిప్) తరపున తెలియజేయునది ఏమనగా, తేదీ ${camp.campDate || ''} న మీ గ్రామం ${camp.village} లో మెగా ఉచిత సర్వీస్ క్యాంప్ నిర్వహించబడుతోంది. మీ ట్రాక్టర్‌ను తీసుకొచ్చి ఉచిత ఆయిల్ చెకప్ మరియు రిపేర్లు చేయించుకోగలరు.${camp.offers ? ` ఆఫర్లు: ${camp.offers}` : ''} సంప్రదించండి: ${camp.contactPhone || 'డీలర్షిప్'}.`
                                  : `Greetings from Sri Gayatri Automotives (Eicher Tractors). We are organizing a Mega Service Camp in your village ${camp.village} on ${camp.campDate || ''}.${camp.offers ? ` Special Offers: ${camp.offers}` : ''} Contact: ${camp.contactPhone || 'Dealership'}.`;

                                setBroadcastData({
                                  villageName: camp.village,
                                  mandalName: camp.mandal,
                                  branchName: camp.branch,
                                  customers: villageCusts,
                                  messageTemplate: defaultMsg
                                });
                              }}
                              className="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-900 text-[10px] font-bold rounded border border-emerald-300 transition cursor-pointer flex items-center gap-1"
                              title={isTe ? 'రైతులకు WhatsApp మెసేజ్ పంపండి' : 'Broadcast WhatsApp message to village farmers'}
                            >
                              <MessageSquare className="w-3 h-3 text-emerald-600" />
                              <span>{villageCusts.length}</span>
                            </button>

                            {/* Maps */}
                            <a
                              href={directionsUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="p-1 text-slate-700 hover:text-blue-900 bg-white border border-slate-300 rounded transition"
                              title="Directions"
                            >
                              <Navigation className="w-3.5 h-3.5" />
                            </a>

                            {/* Edit */}
                            <button
                              type="button"
                              onClick={() => handleEditCamp(camp)}
                              className="p-1 text-slate-700 hover:bg-slate-100 border border-slate-300 rounded transition cursor-pointer"
                              title="Edit"
                            >
                              ✏️
                            </button>

                            {/* Delete */}
                            <button
                              type="button"
                              onClick={() => handleDeleteCamp(camp.id)}
                              className="p-1 text-red-600 hover:bg-red-50 border border-red-200 rounded transition cursor-pointer"
                              title="Delete"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* 3. KRISHNA & NTR DISTRICT ROUTE & DISTANCE MAP VIEW */}
      {/* ========================================================================= */}
      {activeTab === 'district_map' && (
        <KrishnaDistrictRouteMap
          isTe={isTe}
          onSelectCampVillage={(villageName, mandalName, branchName, dealershipCode) => {
            handleOpenNewCampModal({
              village: villageName,
              mandal: mandalName,
              branch: branchName,
              dealershipCode: dealershipCode
            });
          }}
        />
      )}

      {/* ========================================================================= */}
      {/* BULK WHATSAPP MESSAGE & BROADCAST QUEUE MODAL */}
      {/* ========================================================================= */}
      {broadcastData && (
        <BroadcastQueueModal
          villageName={broadcastData.villageName}
          mandalName={broadcastData.mandalName}
          branchName={broadcastData.branchName}
          customers={broadcastData.customers}
          messageTemplate={broadcastData.messageTemplate}
          isTe={isTe}
          onClose={() => setBroadcastData(null)}
          contactedMap={contactedCustomers}
          onMarkContacted={handleMarkContacted}
          showToast={showToast}
        />
      )}

      {/* Floating Toast Notification */}
      {toastMsg && (
        <div className="fixed bottom-5 right-5 z-60 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-2xl flex items-center gap-2.5 border border-slate-700 animate-in fade-in slide-in-from-bottom-3 duration-200 text-xs font-semibold">
          <CheckCircle className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMsg}</span>
          <button
            type="button"
            onClick={() => setToastMsg(null)}
            className="ml-2 text-slate-400 hover:text-white p-0.5 rounded cursor-pointer"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODAL 2: PLAN / EDIT SERVICE CAMP SCHEDULER FORM */}
      {/* ========================================================================= */}
      {isCampModalOpen && editingCamp && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-3 md:p-6 bg-slate-950/60 backdrop-blur-xs">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-2xl max-h-[92vh] flex flex-col overflow-hidden animate-in fade-in zoom-in duration-200">
            {/* Form Header */}
            <div className="p-4 md:p-5 bg-blue-950 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="p-1.5 bg-amber-400 text-slate-950 rounded-md">
                  <Calendar className="w-4 h-4" />
                </span>
                <h3 className="text-lg font-black tracking-tight">
                  {editingCamp.id && serviceCamps.some(c => c.id === editingCamp.id)
                    ? isTe ? 'సర్వీస్ క్యాంప్ ప్లాన్ ఎడిట్ చేయండి' : 'Edit Service Camp Plan'
                    : isTe ? 'కొత్త సర్వీస్ క్యాంప్ ప్లాన్ చేయండి' : 'Schedule New Service Camp'}
                </h3>
              </div>

              <button
                type="button"
                onClick={() => setIsCampModalOpen(false)}
                className="p-1.5 text-white/80 hover:text-white bg-white/10 hover:bg-white/20 rounded-lg transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Form Fields */}
            <form onSubmit={handleSaveCamp} className="flex-1 overflow-y-auto p-5 md:p-6 space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {/* Dealership Code */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'డీలర్‌షిప్ హబ్ కోడ్' : 'Dealership Hub Code'} *
                  </label>
                  <select
                    value={editingCamp.dealershipCode || '4731'}
                    onChange={e => {
                      const code = e.target.value as '4731' | '4732';
                      const defaultBranch = DEALERSHIP_DATA[code].branches[0].name;
                      const defaultMandal = DEALERSHIP_DATA[code].branches[0].mandals[0].name;
                      setEditingCamp(prev => ({
                        ...prev,
                        dealershipCode: code,
                        branch: defaultBranch,
                        mandal: defaultMandal,
                        village: ''
                      }));
                    }}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  >
                    <option value="4731">4731 - Tiruvuru, Nuzvidu, Nandigama</option>
                    <option value="4732">4732 - Gudivada, Poranki, Machilipatnam</option>
                  </select>
                </div>

                {/* Branch */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'మెయిన్ బ్రాంచ్' : 'Main Branch'} *
                  </label>
                  <select
                    value={editingCamp.branch || ''}
                    onChange={e => {
                      const bName = e.target.value;
                      const hub =
                        DEALERSHIP_DATA[editingCamp.dealershipCode as '4731' | '4732'] || DEALERSHIP_DATA['4731'];
                      const branchObj = hub.branches.find(b => b.name === bName) || hub.branches[0];
                      setEditingCamp(prev => ({
                        ...prev,
                        branch: bName,
                        mandal: branchObj.mandals[0]?.name || '',
                        village: ''
                      }));
                    }}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  >
                    {(DEALERSHIP_DATA[editingCamp.dealershipCode as '4731' | '4732'] || DEALERSHIP_DATA['4731']).branches.map(b => (
                      <option key={b.id} value={b.name}>
                        {b.name} ({b.teluguName})
                      </option>
                    ))}
                  </select>
                </div>

                {/* Mandal */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'మండలం' : 'Mandal'} *
                  </label>
                  <select
                    value={editingCamp.mandal || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, mandal: e.target.value, village: '' }))}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  >
                    {(() => {
                      const hub =
                        DEALERSHIP_DATA[editingCamp.dealershipCode as '4731' | '4732'] || DEALERSHIP_DATA['4731'];
                      const branchObj = hub.branches.find(b => b.name === editingCamp.branch) || hub.branches[0];
                      return branchObj.mandals.map(m => (
                        <option key={m.name} value={m.name}>
                          {m.name} {m.teluguName ? `(${m.teluguName})` : ''}
                        </option>
                      ));
                    })()}
                  </select>
                </div>

                {/* Village */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'గ్రామం (Village)' : 'Village'} *
                  </label>
                  <input
                    type="text"
                    list="village-suggestions"
                    value={editingCamp.village || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, village: e.target.value }))}
                    placeholder={isTe ? 'గ్రామం పేరు నమోదు చేయండి లేదా ఎంచుకోండి...' : 'Enter or select village name...'}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                    required
                  />
                  <datalist id="village-suggestions">
                    {(() => {
                      const hub =
                        DEALERSHIP_DATA[editingCamp.dealershipCode as '4731' | '4732'] || DEALERSHIP_DATA['4731'];
                      const branchObj = hub.branches.find(b => b.name === editingCamp.branch) || hub.branches[0];
                      const mandalObj = branchObj.mandals.find(m => m.name === editingCamp.mandal) || branchObj.mandals[0];
                      return (mandalObj?.villages || []).map(v => (
                        <option key={v.name} value={v.name}>
                          {v.name} ({v.distanceKm} km from {editingCamp.branch})
                        </option>
                      ));
                    })()}
                  </datalist>
                </div>

                {/* Camp Date */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'క్యాంప్ తేదీ' : 'Camp Date'} *
                  </label>
                  <input
                    type="date"
                    value={editingCamp.campDate || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, campDate: e.target.value }))}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                    required
                  />
                </div>

                {/* Target Tractors */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'టార్గెట్ ట్రాక్టర్ల సంఖ్య' : 'Target Tractors Count'}
                  </label>
                  <input
                    type="number"
                    min="1"
                    value={editingCamp.targetTractors || '15'}
                    onChange={e => setEditingCamp(prev => ({ ...prev, targetTractors: e.target.value }))}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                </div>

                {/* Venue / Location */}
                <div className="sm:col-span-2">
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'క్యాంప్ ప్రదేశం / వేదిక (Venue)' : 'Camp Venue / Location'}
                  </label>
                  <input
                    type="text"
                    value={editingCamp.venue || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, venue: e.target.value }))}
                    placeholder="e.g. Rythu Bharosa Kendram (RBK) / Gram Panchayat / High School Ground"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                </div>

                {/* Assigned Supervisor */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'బాధ్యత గల సూపర్‌వైజర్' : 'Assigned Supervisor'}
                  </label>
                  <input
                    type="text"
                    list="supervisor-suggestions"
                    value={editingCamp.supervisor || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, supervisor: e.target.value }))}
                    placeholder="Select or enter supervisor name"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                  <datalist id="supervisor-suggestions">
                    {Array.from(new Set(
                      staffList
                        .filter(s => s.role?.toLowerCase().includes('supervisor') || s.role === 'supervisor')
                        .map(s => s.name?.trim())
                        .filter(Boolean)
                    )).map((name, idx) => (
                      <option key={`camp-sup-${name}-${idx}`} value={name} />
                    ))}
                  </datalist>
                </div>

                {/* Assigned Mechanic */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'బాధ్యత గల మెకానిక్ / టెక్నీషియన్' : 'Assigned Mechanic'}
                  </label>
                  <input
                    type="text"
                    list="mechanic-suggestions"
                    value={editingCamp.mechanic || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, mechanic: e.target.value }))}
                    placeholder="Select or enter mechanic name"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                  <datalist id="mechanic-suggestions">
                    {Array.from(new Set(
                      staffList
                        .filter(s => s.role?.toLowerCase().includes('mechanic') || s.role === 'mechanic')
                        .map(s => s.name?.trim())
                        .filter(Boolean)
                    )).map((name, idx) => (
                      <option key={`camp-mech-${name}-${idx}`} value={name} />
                    ))}
                  </datalist>
                </div>

                {/* Status */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'క్యాంప్ స్థితి (Status)' : 'Camp Status'}
                  </label>
                  <select
                    value={editingCamp.status || 'Upcoming'}
                    onChange={e => setEditingCamp(prev => ({ ...prev, status: e.target.value as any }))}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  >
                    <option value="Upcoming">{isTe ? 'Upcoming (రాబోయేది)' : 'Upcoming'}</option>
                    <option value="In Progress">{isTe ? 'In Progress (జరుగుతున్నది)' : 'In Progress'}</option>
                    <option value="Completed">{isTe ? 'Completed (పూర్తయింది)' : 'Completed'}</option>
                    <option value="Cancelled">{isTe ? 'Cancelled (రద్దు)' : 'Cancelled'}</option>
                  </select>
                </div>

                {/* Attended Count */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'హాజరైన / సర్వీస్ చేసిన ట్రాక్టర్లు' : 'Attended / Serviced Count'}
                  </label>
                  <input
                    type="number"
                    min="0"
                    value={editingCamp.attendedCount || '0'}
                    onChange={e => setEditingCamp(prev => ({ ...prev, attendedCount: e.target.value }))}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                </div>

                {/* Village Contact Person & Phone */}
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'గ్రామ సమన్వయకర్త / సర్పంచ్ / రైతు పేరు' : 'Village Contact Person'}
                  </label>
                  <input
                    type="text"
                    value={editingCamp.contactPerson || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, contactPerson: e.target.value }))}
                    placeholder="Name"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'సమన్వయకర్త మొబైల్ నెం.' : 'Contact Mobile No.'}
                  </label>
                  <input
                    type="tel"
                    value={editingCamp.contactPhone || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, contactPhone: e.target.value }))}
                    placeholder="10-digit mobile"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                </div>

                {/* Special Offers & Packages */}
                <div className="sm:col-span-2">
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'ప్రత్యేక ఆఫర్లు & రాయితీలు (Offers & Benefits)' : 'Special Offers & Packages'}
                  </label>
                  <textarea
                    rows={2}
                    value={editingCamp.offers || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, offers: e.target.value }))}
                    placeholder="e.g. Free 50-Point General Inspection, 10% Discount on Eicher Genuine Oil & Filters, Free Greasing..."
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                </div>

                {/* Additional Notes */}
                <div className="sm:col-span-2">
                  <label className="block font-bold text-slate-700 mb-1">
                    {isTe ? 'అదనపు గమనికలు (Notes)' : 'Additional Notes'}
                  </label>
                  <textarea
                    rows={2}
                    value={editingCamp.notes || ''}
                    onChange={e => setEditingCamp(prev => ({ ...prev, notes: e.target.value }))}
                    placeholder="Any special remarks..."
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:bg-white focus:ring-2 focus:ring-blue-900"
                  />
                </div>
              </div>

              {/* Form Actions */}
              <div className="pt-4 border-t border-slate-200 flex items-center justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setIsCampModalOpen(false)}
                  className="px-4 py-2 text-slate-700 bg-slate-100 hover:bg-slate-200 font-bold rounded-lg transition cursor-pointer"
                >
                  {isTe ? 'రద్దు చేయండి' : 'Cancel'}
                </button>
                <button
                  type="submit"
                  disabled={savingCamp}
                  className="px-5 py-2 bg-blue-900 hover:bg-blue-800 text-white font-bold rounded-lg shadow-sm transition cursor-pointer flex items-center gap-1.5"
                >
                  <Check className="w-4 h-4" />
                  <span>
                    {savingCamp
                      ? isTe ? 'సేవ్ అవుతోంది...' : 'Saving...'
                      : isTe ? 'క్యాంప్ ప్లాన్ సేవ్ చేయండి' : 'Save Camp Plan'}
                  </span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default ServiceCampPlanning;
