import React, { useState } from 'react';
import { 
  Building2, 
  MapPin, 
  ChevronRight, 
  ChevronDown, 
  Folder, 
  FolderOpen,
  Plus,
  Trash2,
  Edit2,
  Save,
  X
} from 'lucide-react';

interface GeoSetupProps {
  dealershipData: any;
  isTe: boolean;
  onUpdateData: (newData: any) => void;
}

export const GeoSetup: React.FC<GeoSetupProps> = ({ dealershipData, isTe, onUpdateData }) => {
  const [expandedHubs, setExpandedHubs] = useState<Record<string, boolean>>({ '4731': true });
  const [expandedBranches, setExpandedBranches] = useState<Record<string, boolean>>({});
  const [expandedMandals, setExpandedMandals] = useState<Record<string, boolean>>({});

  const toggleHub = (code: string) => setExpandedHubs(p => ({ ...p, [code]: !p[code] }));
  const toggleBranch = (id: string) => setExpandedBranches(p => ({ ...p, [id]: !p[id] }));
  const toggleMandal = (key: string) => setExpandedMandals(p => ({ ...p, [key]: !p[key] }));

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-full">
      <div className="bg-slate-900 text-white px-4 py-3 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <Building2 className="w-5 h-5 text-amber-400" />
          <h2 className="font-black text-sm uppercase tracking-wider">
            {isTe ? 'జియోగ్రాఫిక్ సెటప్ (Branches, Mandals, Villages)' : 'Geographic Setup'}
          </h2>
        </div>
        <div className="text-[10px] bg-white/10 px-2 py-1 rounded font-bold">
          MASTER DIRECTORY CONTROL
        </div>
      </div>

      <div className="flex-1 overflow-y-auto p-4 bg-slate-50">
        <div className="max-w-4xl mx-auto space-y-4">
          {Object.entries(dealershipData).map(([code, hub]: [string, any]) => {
            const isHubExpanded = expandedHubs[code];
            return (
              <div key={code} className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
                {/* HUB ROW */}
                <div 
                  onClick={() => toggleHub(code)}
                  className="px-4 py-3 flex items-center justify-between bg-slate-900 text-white cursor-pointer"
                >
                  <div className="flex items-center gap-3">
                    {isHubExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
                    <span className="font-black text-sm">HUB {code}: {isTe ? hub.teluguName : hub.name}</span>
                  </div>
                  <span className="text-[10px] font-bold opacity-70">{hub.branches.length} Branches</span>
                </div>

                {isHubExpanded && (
                  <div className="p-2 space-y-2">
                    {hub.branches.map((branch: any) => {
                      const isBranchExpanded = expandedBranches[branch.id];
                      return (
                        <div key={branch.id} className="border border-slate-100 rounded-lg overflow-hidden bg-white">
                          <div 
                            onClick={() => toggleBranch(branch.id)}
                            className="px-4 py-2.5 flex items-center justify-between bg-slate-50 hover:bg-slate-100 cursor-pointer transition"
                          >
                            <div className="flex items-center gap-3">
                              {isBranchExpanded ? <FolderOpen className="w-4 h-4 text-blue-900" /> : <Folder className="w-4 h-4 text-slate-400" />}
                              <span className="font-bold text-xs text-slate-800">{branch.name}</span>
                            </div>
                            <div className="flex items-center gap-3">
                              <span className="text-[10px] text-slate-500 font-bold">{branch.mandals.length} Mandals</span>
                              <button className="p-1 hover:bg-blue-100 text-blue-600 rounded"><Edit2 className="w-3.5 h-3.5" /></button>
                            </div>
                          </div>

                          {isBranchExpanded && (
                            <div className="pl-8 pr-4 py-2 space-y-2 bg-slate-50/50">
                              {branch.mandals.map((mandal: any) => {
                                const mKey = `${branch.id}-${mandal.name}`;
                                const isMExpand = expandedMandals[mKey];
                                return (
                                  <div key={mandal.name} className="border border-slate-200 rounded-md bg-white overflow-hidden shadow-3xs">
                                    <div 
                                      onClick={() => toggleMandal(mKey)}
                                      className="px-3 py-2 flex items-center justify-between hover:bg-slate-50 cursor-pointer"
                                    >
                                      <div className="flex items-center gap-2">
                                        {isMExpand ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />}
                                        <span className="font-bold text-[11px] text-slate-700">{mandal.name} Mandal</span>
                                      </div>
                                      <div className="flex items-center gap-2">
                                        <span className="text-[10px] text-slate-400 font-bold">{mandal.villages.length} Villages</span>
                                        <button className="p-1 hover:bg-emerald-100 text-emerald-600 rounded"><Plus className="w-3 h-3" /></button>
                                      </div>
                                    </div>

                                    {isMExpand && (
                                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 p-3 border-t border-slate-100 bg-slate-50/30">
                                        {mandal.villages.map((village: any) => (
                                          <div key={village.name} className="flex items-center justify-between p-2 bg-white border border-slate-200 rounded-md shadow-3xs group">
                                            <div className="flex flex-col">
                                              <span className="text-[11px] font-bold text-slate-800">{village.name}</span>
                                              <span className="text-[9px] text-slate-400 font-medium">{village.distanceKm} km away</span>
                                            </div>
                                            <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition">
                                              <button className="p-1 hover:bg-slate-100 text-slate-400 rounded"><Edit2 className="w-2.5 h-2.5" /></button>
                                              <button className="p-1 hover:bg-rose-50 text-rose-500 rounded"><Trash2 className="w-2.5 h-2.5" /></button>
                                            </div>
                                          </div>
                                        ))}
                                        <button className="flex items-center justify-center gap-2 p-2 border-2 border-dashed border-slate-200 rounded-md hover:border-blue-300 hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition group">
                                          <Plus className="w-3.5 h-3.5" />
                                          <span className="text-[10px] font-bold">Add Village</span>
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                );
                              })}
                              <button className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 rounded-md hover:border-blue-300 text-slate-500 hover:text-blue-600 transition w-full">
                                <Plus className="w-4 h-4" />
                                <span className="text-xs font-bold">Add New Mandal to {branch.name}</span>
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                    <button className="flex items-center gap-2 px-4 py-3 bg-slate-50 border-2 border-dashed border-slate-200 rounded-xl hover:border-blue-400 hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition w-full">
                      <Plus className="w-5 h-5" />
                      <span className="text-sm font-bold">Add New Branch to Hub {code}</span>
                    </button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
      
      <div className="bg-slate-100 p-4 border-t border-slate-200 flex justify-between items-center">
        <p className="text-[10px] text-slate-500 font-medium max-w-md italic">
          {isTe ? '* ఈ సెటప్‌లో చేసే మార్పులు కేవలం సేవా క్యాంప్ ప్లానింగ్ పేజీలో మాత్రమే ప్రతిబింబిస్తాయి.' : '* Changes made here affect the Geographic structure used in Service Camp Planning.'}
        </p>
        <button className="px-6 py-2 bg-blue-900 text-white font-black text-xs rounded-lg shadow-md hover:bg-blue-950 transition">
          {isTe ? 'మార్పులను సేవ్ చేయండి' : 'SAVE MASTER CHANGES'}
        </button>
      </div>
    </div>
  );
};
