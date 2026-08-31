// Real PostgreSQL REST API client for Sri Balaji Eicher Tractors

export const sqlApi = {
  // Customers
  fetchCustomers: async () => {
    try {
      const res = await fetch('/api/customers');
      const data = await res.json();
      if (data.success && Array.isArray(data.data)) {
        return data.data.map((r: any) => {
          let fullData = {};
          let followupHistory = [];
          try { fullData = typeof r.full_data === 'string' ? JSON.parse(r.full_data) : (r.full_data || {}); } catch {}
          try { followupHistory = typeof r.followup_history === 'string' ? JSON.parse(r.followup_history) : (r.followup_history || []); } catch {}
          return {
            ...fullData,
            id: r.id,
            chassisKey: r.chassis_key,
            chassisNo: r.chassis_no,
            custName: r.cust_name,
            fatherName: r.father_name,
            custAddr: r.cust_addr,
            village: r.village,
            mandal: r.mandal,
            ownerMob: r.owner_mob,
            driverMob: r.driver_mob,
            regdNo: r.regd_no,
            engineNo: r.engine_no,
            tractorModel: r.tractor_model,
            dateOfDelivery: r.date_of_delivery,
            followupHistory,
          };
        });
      }
      return [];
    } catch (e) {
      console.error('Error fetching customers from PostgreSQL:', e);
      return [];
    }
  },
  getCustomers: async () => sqlApi.fetchCustomers(),

  saveCustomer: async (customer: any) => {
    try {
      const res = await fetch('/api/customers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(customer)
      });
      return await res.json();
    } catch (e) {
      console.error('Error saving customer to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  bulkUpsertCustomers: async (rows: any[], replaceAll = false) => {
    try {
      const res = await fetch('/api/customers/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows, replaceAll })
      });
      return await res.json();
    } catch (e) {
      console.error('Error bulk saving customers to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },
  saveCustomersBulk: async (rows: any[], replaceAll = false) => sqlApi.bulkUpsertCustomers(rows, replaceAll),

  deleteAllCustomers: async () => {
    try {
      const res = await fetch('/api/database/delete-all-customers', {
        method: 'POST',
      });
      return await res.json();
    } catch (e) {
      console.error('Error deleting all customers from PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  // Spares
  fetchSpares: async () => {
    try {
      const res = await fetch('/api/spares');
      const data = await res.json();
      if (data.success && Array.isArray(data.data)) {
        return data.data.map((r: any) => {
          let fullData = {};
          try { fullData = typeof r.full_data === 'string' ? JSON.parse(r.full_data) : (r.full_data || {}); } catch {}
          return {
            ...fullData,
            id: r.id,
            partKey: r.part_key,
            partNo: r.part_no,
            partDesc: r.part_desc,
            mrp: r.mrp,
            category: r.category
          };
        });
      }
      return [];
    } catch (e) {
      console.error('Error fetching spares from PostgreSQL:', e);
      return [];
    }
  },
  getSpares: async () => sqlApi.fetchSpares(),

  saveSpare: async (spare: any) => {
    try {
      const res = await fetch('/api/spares/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows: [spare], replaceAll: false })
      });
      return await res.json();
    } catch (e) {
      console.error('Error saving spare to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  bulkUpsertSpares: async (rows: any[], replaceAll = false) => {
    try {
      const res = await fetch('/api/spares/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows, replaceAll })
      });
      return await res.json();
    } catch (e) {
      console.error('Error bulk saving spares to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },
  saveSparesBulk: async (rows: any[], replaceAll = false) => sqlApi.bulkUpsertSpares(rows, replaceAll),

  // Jobcards
  fetchJobcards: async () => {
    try {
      const res = await fetch('/api/jobcards');
      const data = await res.json();
      if (data.success && Array.isArray(data.data)) {
        return data.data.map((r: any) => {
          let checkpoints = [];
          let repairRows = [];
          let partRows = [];
          try { checkpoints = typeof r.checkpoints === 'string' ? JSON.parse(r.checkpoints) : (r.checkpoints || []); } catch {}
          try { repairRows = typeof r.repair_rows === 'string' ? JSON.parse(r.repair_rows) : (r.repair_rows || []); } catch {}
          try { partRows = typeof r.part_rows === 'string' ? JSON.parse(r.part_rows) : (r.part_rows || []); } catch {}

          return {
            id: r.id,
            jobNo: r.job_no,
            onlineJobCardNo: r.online_job_card_no,
            jobDate: r.job_date,
            dateTimeIn: r.date_time_in,
            dateTimeOut: r.date_time_out,
            expectedRepairTime: r.expected_repair_time,
            status: r.status,
            custName: r.cust_name,
            fatherName: r.father_name,
            custAddr: r.cust_addr,
            village: r.village,
            mandal: r.mandal,
            ownerMob: r.owner_mob,
            driverMob: r.driver_mob,
            regdNo: r.regd_no,
            chassisNo: r.chassis_no,
            engineNo: r.engine_no,
            model: r.model,
            modelType: r.model_type,
            serialNo: r.serial_no,
            hourMeter: r.hour_meter,
            serviceType: r.service_type,
            freeServiceList: r.free_service_list,
            extraRepairs: r.extra_repairs,
            mechanic: r.mechanic,
            mechanicName: r.mechanic,
            wsIncharge: r.ws_incharge,
            supervisor: r.ws_incharge,
            serviceLocation: r.service_location,
            billNo: r.bill_no,
            reasonsForAnalysis: r.reasons_for_analysis,
            telecalling: r.telecalling,
            warrantyOverride: r.warranty_override,
            totalLabour: r.total_labour,
            warrantyMaterial: r.warranty_material,
            nonWarrantyMaterial: r.non_warranty_material,
            gTotal: r.g_total,
            actualClosedDate: r.actual_closed_date,
            checkpoints,
            repairRows,
            partRows,
            createdBy: r.created_by,
            createdByEmail: r.created_by_email,
            createdAt: r.created_at,
          };
        });
      }
      return [];
    } catch (e) {
      console.error('Error fetching jobcards from PostgreSQL:', e);
      return [];
    }
  },
  getJobCards: async () => sqlApi.fetchJobcards(),

  saveJobcard: async (card: any) => {
    try {
      const res = await fetch('/api/jobcards', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(card)
      });
      return await res.json();
    } catch (e) {
      console.error('Error saving jobcard to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },
  saveJobCard: async (card: any) => sqlApi.saveJobcard(card),

  deleteJobcard: async (id: string) => {
    try {
      const res = await fetch(`/api/jobcards/${encodeURIComponent(id)}`, {
        method: 'DELETE'
      });
      return await res.json();
    } catch (e) {
      console.error('Error deleting jobcard from PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },
  deleteJobCard: async (id: string) => sqlApi.deleteJobcard(id),

  bulkUpsertJobcards: async (cards: any[], replaceAll = false) => {
    try {
      const res = await fetch('/api/jobcards/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ cards, replaceAll })
      });
      return await res.json();
    } catch (e) {
      console.error('Error bulk saving jobcards to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },
  saveJobCardsBulk: async (cards: any[], replaceAll = false) => sqlApi.bulkUpsertJobcards(cards, replaceAll),

  bulkDeleteJobcards: async (ids: string[]) => {
    try {
      const res = await fetch('/api/database/delete-jobcards', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids })
      });
      return await res.json();
    } catch (e) {
      console.error('Error bulk deleting jobcards from PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  // Complaints
  fetchComplaints: async () => {
    try {
      const res = await fetch('/api/complaints');
      const data = await res.json();
      if (data.success && Array.isArray(data.data)) {
        return data.data.map((r: any) => ({
          id: r.id,
          complaintNo: r.complaint_no,
          date: r.date,
          customerName: r.customer_name,
          phone: r.phone,
          village: r.village,
          mandal: r.mandal,
          tractorModel: r.tractor_model,
          chassisNo: r.chassis_no,
          hours: r.hours,
          complaintDetails: r.complaint_details,
          mechanic: r.mechanic,
          status: r.status,
          jobCardNo: r.job_card_no,
          closureDate: r.closure_date,
          remarks: r.remarks,
          createdAt: r.created_at
        }));
      }
      return [];
    } catch (e) {
      console.error('Error fetching complaints from PostgreSQL:', e);
      return [];
    }
  },
  getComplaints: async () => sqlApi.fetchComplaints(),

  saveComplaint: async (complaint: any) => {
    try {
      const res = await fetch('/api/complaints', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(complaint)
      });
      return await res.json();
    } catch (e) {
      console.error('Error saving complaint to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  deleteComplaint: async (id: string) => {
    try {
      const res = await fetch(`/api/complaints/${encodeURIComponent(id)}`, {
        method: 'DELETE'
      });
      return await res.json();
    } catch (e) {
      console.error('Error deleting complaint from PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  bulkUpsertComplaints: async (complaints: any[], replaceAll = false) => {
    try {
      const res = await fetch('/api/complaints/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ complaints, replaceAll })
      });
      return await res.json();
    } catch (e) {
      console.error('Error bulk saving complaints to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },
  saveComplaintsBulk: async (complaints: any[], replaceAll = false) => sqlApi.bulkUpsertComplaints(complaints, replaceAll),

  // Staff
  fetchStaff: async () => {
    try {
      const res = await fetch('/api/staff');
      const data = await res.json();
      if (data.success && Array.isArray(data.data)) {
        return data.data.map((r: any) => ({
          id: r.id,
          name: r.name,
          role: r.role,
          phone: r.phone,
          active: r.active === 'true' || r.active === true,
          assignedSupervisor: r.assigned_supervisor,
          createdAt: r.created_at
        }));
      }
      return [];
    } catch (e) {
      console.error('Error fetching staff from PostgreSQL:', e);
      return [];
    }
  },
  getStaff: async () => sqlApi.fetchStaff(),

  saveStaff: async (staff: any) => {
    try {
      const res = await fetch('/api/staff', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(staff)
      });
      return await res.json();
    } catch (e) {
      console.error('Error saving staff to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  deleteStaff: async (id: string) => {
    try {
      const res = await fetch(`/api/staff/${encodeURIComponent(id)}`, {
        method: 'DELETE'
      });
      return await res.json();
    } catch (e) {
      console.error('Error deleting staff from PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  bulkUpsertStaff: async (staff: any[], replaceAll = false) => {
    try {
      const res = await fetch('/api/staff/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ staff, replaceAll })
      });
      return await res.json();
    } catch (e) {
      console.error('Error bulk saving staff to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },
  saveStaffBulk: async (staff: any[], replaceAll = false) => sqlApi.bulkUpsertStaff(staff, replaceAll),

  // Attendance
  fetchAttendance: async () => {
    try {
      const res = await fetch('/api/attendance');
      const data = await res.json();
      if (data.success && data.data) {
        return data.data;
      }
      return {};
    } catch (e) {
      console.error('Error fetching attendance from PostgreSQL:', e);
      return {};
    }
  },
  getAttendance: async () => sqlApi.fetchAttendance(),

  saveAttendance: async (date: string, records: any) => {
    try {
      const res = await fetch('/api/attendance', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date, records })
      });
      return await res.json();
    } catch (e) {
      console.error('Error saving attendance to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  // App Settings
  fetchSettings: async () => {
    try {
      const res = await fetch('/api/settings');
      const data = await res.json();
      if (data.success && data.data) {
        return data.data;
      }
      return {};
    } catch (e) {
      console.error('Error fetching settings from PostgreSQL:', e);
      return {};
    }
  },
  getSettings: async () => sqlApi.fetchSettings(),

  saveSettings: async (key: string, value: string) => {
    try {
      const res = await fetch('/api/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key, value })
      });
      return await res.json();
    } catch (e) {
      console.error('Error saving setting to PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  },

  // Master Backup
  fetchMasterBackup: async () => {
    try {
      const res = await fetch('/api/master-backup');
      return await res.json();
    } catch (e) {
      console.error('Error fetching master backup from PostgreSQL:', e);
      return { success: false, error: String(e) };
    }
  }
};

export const sqlService = sqlApi;
