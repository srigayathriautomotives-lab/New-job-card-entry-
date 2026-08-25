import express from 'express';
import cors from 'cors';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import pg from 'pg';

const { Pool } = pg;

const app = express();
const PORT = Number(process.env.PORT || 3000);

app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

let pool: any = null;

function getPool(): any {
  if (!pool) {
    try {
      if (process.env.DATABASE_URL) {
        pool = new Pool({
          connectionString: process.env.DATABASE_URL,
        });
      } else {
        throw new Error('DATABASE_URL is missing');
      }
    } catch {
      console.warn('DB not connected — mock active');
      pool = {
        query: async () => ({ rows: [] }),
        connect: async () => ({ query: async () => ({ rows: [] }), release: () => {} })
      };
    }
  }
  return pool;
}

// Clear database API
app.post('/api/database/clear', async (req, res) => {
  const { password } = req.body;
  if (password !== 'AdminClear123') { // Simple placeholder check
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }

  const client = await getPool().connect();
  console.log('Database clearing initiated on configured PostgreSQL database');
  try {
    await client.query('BEGIN');
    await client.query('SET session_replication_role = replica;');
    await client.query('TRUNCATE TABLE customers, jobcards, complaints, spares, staff, staff_attendance, app_settings RESTART IDENTITY CASCADE');
    await client.query('SET session_replication_role = origin;');
    await client.query('COMMIT');
    console.log('Database cleared successfully.');
    res.json({ success: true });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error clearing database:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// Bulk delete jobcards API
app.post('/api/database/delete-jobcards', async (req, res) => {
  const { ids } = req.body;
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ success: false, error: 'No IDs provided' });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    // Using unnest to handle array of IDs efficiently
    await client.query('DELETE FROM jobcards WHERE id = ANY($1)', [ids]);
    await client.query('COMMIT');
    res.json({ success: true });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error deleting jobcards:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// Health check API
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// =======================
// CUSTOMERS API
// =======================

// Get all customers
app.get('/api/customers', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM customers ORDER BY id ASC');
    res.json({ success: true, data: result.rows });
  } catch (error: any) {
    console.error('Error fetching customers:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Bulk upsert/replace customers (from Excel/Backup upload)
app.post('/api/customers/bulk', async (req, res) => {
  const { rows, replaceAll } = req.body;
  if (!Array.isArray(rows)) {
    return res.status(400).json({ success: false, error: 'Rows array required' });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');

    if (replaceAll) {
      await client.query('DELETE FROM customers');
    }

    const insertQuery = `
      INSERT INTO customers (
        chassis_key, chassis_no, cust_name, father_name, cust_addr, village, mandal,
        owner_mob, driver_mob, regd_no, engine_no, tractor_model, date_of_delivery,
        followup_history, full_data, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, NOW())
      ON CONFLICT (chassis_key) DO UPDATE SET
        chassis_no = EXCLUDED.chassis_no,
        cust_name = EXCLUDED.cust_name,
        father_name = EXCLUDED.father_name,
        cust_addr = EXCLUDED.cust_addr,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        owner_mob = EXCLUDED.owner_mob,
        driver_mob = EXCLUDED.driver_mob,
        regd_no = EXCLUDED.regd_no,
        engine_no = EXCLUDED.engine_no,
        tractor_model = EXCLUDED.tractor_model,
        date_of_delivery = EXCLUDED.date_of_delivery,
        followup_history = EXCLUDED.followup_history,
        full_data = EXCLUDED.full_data,
        updated_at = NOW()
    `;

    for (const r of rows) {
      const chassisKey = r.chassisKey || r.chassis_key || r.key || (r.chassisNo ? String(r.chassisNo).toLowerCase().replace(/[^a-z0-9]/g, '') : null);
      if (!chassisKey) continue;

      const chassisNo = r.chassisNo || r.chassis_no || '';
      const custName = r.custName || r.cust_name || r['cutomer name'] || r['Customer Name'] || '';
      const fatherName = r.fatherName || r.father_name || r.father || '';
      const custAddr = r.custAddr || r.cust_addr || r.address || '';
      const village = r.village || '';
      const mandal = r.mandal || '';
      const ownerMob = r.ownerMob || r.owner_mob || r.phone || r['ph no'] || '';
      const driverMob = r.driverMob || r.driver_mob || '';
      const regdNo = r.regdNo || r.regd_no || '';
      const engineNo = r.engineNo || r.engine_no || '';
      const tractorModel = r.tractorModel || r.tractor_model || r.model || '';
      const dateOfDelivery = r.dateOfDelivery || r.date_of_delivery || '';
      const followupHistory = JSON.stringify(r.followupHistory || r.followup_history || []);
      const fullData = JSON.stringify(r.fullData || r.full_data || r);

      await client.query(insertQuery, [
        chassisKey, chassisNo, custName, fatherName, custAddr, village, mandal,
        ownerMob, driverMob, regdNo, engineNo, tractorModel, dateOfDelivery,
        followupHistory, fullData
      ]);
    }

    await client.query('COMMIT');
    res.json({ success: true, count: rows.length });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error saving customers bulk:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// Upsert single customer
app.post('/api/customers', async (req, res) => {
  const r = req.body;
  const chassisKey = r.chassisKey || r.chassis_key || (r.chassisNo ? String(r.chassisNo).toLowerCase().replace(/[^a-z0-9]/g, '') : null);
  if (!chassisKey) {
    return res.status(400).json({ success: false, error: 'Valid chassis number is required' });
  }

  try {
    const client = getPool();
    const query = `
      INSERT INTO customers (
        chassis_key, chassis_no, cust_name, father_name, cust_addr, village, mandal,
        owner_mob, driver_mob, regd_no, engine_no, tractor_model, date_of_delivery,
        followup_history, full_data, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, NOW())
      ON CONFLICT (chassis_key) DO UPDATE SET
        chassis_no = EXCLUDED.chassis_no,
        cust_name = EXCLUDED.cust_name,
        father_name = EXCLUDED.father_name,
        cust_addr = EXCLUDED.cust_addr,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        owner_mob = EXCLUDED.owner_mob,
        driver_mob = EXCLUDED.driver_mob,
        regd_no = EXCLUDED.regd_no,
        engine_no = EXCLUDED.engine_no,
        tractor_model = EXCLUDED.tractor_model,
        date_of_delivery = EXCLUDED.date_of_delivery,
        followup_history = EXCLUDED.followup_history,
        full_data = EXCLUDED.full_data,
        updated_at = NOW()
      RETURNING *;
    `;
    const result = await client.query(query, [
      chassisKey,
      r.chassisNo || '',
      r.custName || '',
      r.fatherName || '',
      r.custAddr || '',
      r.village || '',
      r.mandal || '',
      r.ownerMob || '',
      r.driverMob || '',
      r.regdNo || '',
      r.engineNo || '',
      r.tractorModel || '',
      r.dateOfDelivery || '',
      JSON.stringify(r.followupHistory || []),
      JSON.stringify(r.fullData || r)
    ]);
    res.json({ success: true, data: result.rows[0] });
  } catch (error: any) {
    console.error('Error saving customer:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// =======================
// SPARES API
// =======================

// Get all spares
app.get('/api/spares', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM spares ORDER BY id ASC');
    res.json({ success: true, data: result.rows });
  } catch (error: any) {
    console.error('Error fetching spares:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Bulk upsert/replace spares (from Excel upload)
app.post('/api/spares/bulk', async (req, res) => {
  const { rows, replaceAll } = req.body;
  if (!Array.isArray(rows)) {
    return res.status(400).json({ success: false, error: 'Rows array required' });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');

    if (replaceAll) {
      await client.query('DELETE FROM spares');
    }

    const insertQuery = `
      INSERT INTO spares (
        part_key, part_no, part_desc, mrp, category, full_data, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, NOW())
      ON CONFLICT (part_key) DO UPDATE SET
        part_no = EXCLUDED.part_no,
        part_desc = EXCLUDED.part_desc,
        mrp = EXCLUDED.mrp,
        category = EXCLUDED.category,
        full_data = EXCLUDED.full_data,
        updated_at = NOW()
    `;

    for (const r of rows) {
      const partKey = r.partKey || r.part_key || (r.partNo ? String(r.partNo).toLowerCase().replace(/[^a-z0-9]/g, '') : null);
      if (!partKey) continue;

      const partNo = r.partNo || r.part_no || '';
      const partDesc = r.partDesc || r.part_desc || r.desc || r.description || '';
      const mrp = r.mrp != null ? String(r.mrp) : '';
      const category = r.category || '';
      const fullData = JSON.stringify(r.fullData || r.full_data || r);

      await client.query(insertQuery, [partKey, partNo, partDesc, mrp, category, fullData]);
    }

    await client.query('COMMIT');
    res.json({ success: true, count: rows.length });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error saving spares bulk:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// =======================
// JOB CARDS API
// =======================

// Get all job cards
app.get('/api/jobcards', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM jobcards ORDER BY created_at DESC');
    res.json({ success: true, data: result.rows });
  } catch (error: any) {
    console.error('Error fetching jobcards:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Save/Update single job card
app.post('/api/jobcards', async (req, res) => {
  const card = req.body;
  if (!card || !card.id) {
    return res.status(400).json({ success: false, error: 'Job card ID is required' });
  }

  try {
    const client = getPool();
    const query = `
      INSERT INTO jobcards (
        id, job_no, online_job_card_no, job_date, date_time_in, date_time_out,
        expected_repair_time, status, cust_name, father_name, cust_addr,
        village, mandal, owner_mob, driver_mob, regd_no, chassis_no, engine_no,
        model, model_type, serial_no, hour_meter, service_type, free_service_list,
        extra_repairs, mechanic, ws_incharge, service_location, bill_no,
        reasons_for_analysis, telecalling, warranty_override, total_labour,
        warranty_material, non_warranty_material, g_total, actual_closed_date,
        checkpoints, repair_rows, part_rows, created_by, created_by_email, created_at, updated_at
      ) VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16,
        $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, $28, $29, $30,
        $31, $32, $33, $34, $35, $36, $37, $38, $39, $40, $41, $42, $43, NOW()
      )
      ON CONFLICT (id) DO UPDATE SET
        job_no = EXCLUDED.job_no,
        online_job_card_no = EXCLUDED.online_job_card_no,
        job_date = EXCLUDED.job_date,
        date_time_in = EXCLUDED.date_time_in,
        date_time_out = EXCLUDED.date_time_out,
        expected_repair_time = EXCLUDED.expected_repair_time,
        status = EXCLUDED.status,
        cust_name = EXCLUDED.cust_name,
        father_name = EXCLUDED.father_name,
        cust_addr = EXCLUDED.cust_addr,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        owner_mob = EXCLUDED.owner_mob,
        driver_mob = EXCLUDED.driver_mob,
        regd_no = EXCLUDED.regd_no,
        chassis_no = EXCLUDED.chassis_no,
        engine_no = EXCLUDED.engine_no,
        model = EXCLUDED.model,
        model_type = EXCLUDED.model_type,
        serial_no = EXCLUDED.serial_no,
        hour_meter = EXCLUDED.hour_meter,
        service_type = EXCLUDED.service_type,
        free_service_list = EXCLUDED.free_service_list,
        extra_repairs = EXCLUDED.extra_repairs,
        mechanic = EXCLUDED.mechanic,
        ws_incharge = EXCLUDED.ws_incharge,
        service_location = EXCLUDED.service_location,
        bill_no = EXCLUDED.bill_no,
        reasons_for_analysis = EXCLUDED.reasons_for_analysis,
        telecalling = EXCLUDED.telecalling,
        warranty_override = EXCLUDED.warranty_override,
        total_labour = EXCLUDED.total_labour,
        warranty_material = EXCLUDED.warranty_material,
        non_warranty_material = EXCLUDED.non_warranty_material,
        g_total = EXCLUDED.g_total,
        actual_closed_date = EXCLUDED.actual_closed_date,
        checkpoints = EXCLUDED.checkpoints,
        repair_rows = EXCLUDED.repair_rows,
        part_rows = EXCLUDED.part_rows,
        created_by = EXCLUDED.created_by,
        created_by_email = EXCLUDED.created_by_email,
        updated_at = NOW()
      RETURNING *;
    `;

    const result = await client.query(query, [
      card.id,
      card.jobNo || card.job_no || '',
      card.onlineJobCardNo || card.online_job_card_no || '',
      card.jobDate || card.job_date || '',
      card.dateTimeIn || card.date_time_in || '',
      card.dateTimeOut || card.date_time_out || '',
      card.expectedRepairTime || card.expected_repair_time || '',
      card.status || 'Open',
      card.custName || card.cust_name || '',
      card.fatherName || card.father_name || '',
      card.custAddr || card.cust_addr || '',
      card.village || '',
      card.mandal || '',
      card.ownerMob || card.owner_mob || '',
      card.driverMob || card.driver_mob || '',
      card.regdNo || card.regd_no || '',
      card.chassisNo || card.chassis_no || '',
      card.engineNo || card.engine_no || '',
      card.model || '',
      card.modelType || card.model_type || '',
      card.serialNo || card.serial_no || '',
      card.hourMeter || card.hour_meter || '',
      card.serviceType || card.service_type || '',
      card.freeServiceList || card.free_service_list || '',
      card.extraRepairs || card.extra_repairs || '',
      card.mechanic || '',
      card.wsIncharge || card.ws_incharge || '',
      card.serviceLocation || card.service_location || '',
      card.billNo || card.bill_no || '',
      card.reasonsForAnalysis || card.reasons_for_analysis || '',
      card.telecalling || '',
      card.warrantyOverride || card.warranty_override || 'auto',
      card.totalLabour || card.total_labour || '',
      card.warrantyMaterial || card.warranty_material || '',
      card.nonWarrantyMaterial || card.non_warranty_material || '',
      card.gTotal || card.g_total || '',
      card.actualClosedDate || card.actual_closed_date || '',
      JSON.stringify(card.checkpoints || []),
      JSON.stringify(card.repairRows || card.repair_rows || []),
      JSON.stringify(card.partRows || card.part_rows || []),
      card.createdBy || card.created_by || '',
      card.createdByEmail || card.created_by_email || '',
      card.createdAt || card.created_at || new Date().toISOString()
    ]);

    res.json({ success: true, data: result.rows[0] });
  } catch (error: any) {
    console.error('Error saving job card:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Delete job card
app.delete('/api/jobcards/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const client = getPool();
    await client.query('DELETE FROM jobcards WHERE id = $1', [id]);
    res.json({ success: true });
  } catch (error: any) {
    console.error('Error deleting jobcard:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Bulk upsert job cards
app.post('/api/jobcards/bulk', async (req, res) => {
  const { cards, replaceAll } = req.body;
  if (!Array.isArray(cards)) {
    return res.status(400).json({ success: false, error: 'Cards array required' });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    if (replaceAll) {
      await client.query('DELETE FROM jobcards');
    }

    const query = `
      INSERT INTO jobcards (
        id, job_no, online_job_card_no, job_date, date_time_in, date_time_out,
        expected_repair_time, status, cust_name, father_name, cust_addr,
        village, mandal, owner_mob, driver_mob, regd_no, chassis_no, engine_no,
        model, model_type, serial_no, hour_meter, service_type, free_service_list,
        extra_repairs, mechanic, ws_incharge, service_location, bill_no,
        reasons_for_analysis, telecalling, warranty_override, total_labour,
        warranty_material, non_warranty_material, g_total, actual_closed_date,
        checkpoints, repair_rows, part_rows, created_by, created_by_email, created_at, updated_at
      ) VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16,
        $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, $28, $29, $30,
        $31, $32, $33, $34, $35, $36, $37, $38, $39, $40, $41, $42, $43, NOW()
      )
      ON CONFLICT (id) DO UPDATE SET
        job_no = EXCLUDED.job_no,
        online_job_card_no = EXCLUDED.online_job_card_no,
        job_date = EXCLUDED.job_date,
        date_time_in = EXCLUDED.date_time_in,
        date_time_out = EXCLUDED.date_time_out,
        expected_repair_time = EXCLUDED.expected_repair_time,
        status = EXCLUDED.status,
        cust_name = EXCLUDED.cust_name,
        father_name = EXCLUDED.father_name,
        cust_addr = EXCLUDED.cust_addr,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        owner_mob = EXCLUDED.owner_mob,
        driver_mob = EXCLUDED.driver_mob,
        regd_no = EXCLUDED.regd_no,
        chassis_no = EXCLUDED.chassis_no,
        engine_no = EXCLUDED.engine_no,
        model = EXCLUDED.model,
        model_type = EXCLUDED.model_type,
        serial_no = EXCLUDED.serial_no,
        hour_meter = EXCLUDED.hour_meter,
        service_type = EXCLUDED.service_type,
        free_service_list = EXCLUDED.free_service_list,
        extra_repairs = EXCLUDED.extra_repairs,
        mechanic = EXCLUDED.mechanic,
        ws_incharge = EXCLUDED.ws_incharge,
        service_location = EXCLUDED.service_location,
        bill_no = EXCLUDED.bill_no,
        reasons_for_analysis = EXCLUDED.reasons_for_analysis,
        telecalling = EXCLUDED.telecalling,
        warranty_override = EXCLUDED.warranty_override,
        total_labour = EXCLUDED.total_labour,
        warranty_material = EXCLUDED.warranty_material,
        non_warranty_material = EXCLUDED.non_warranty_material,
        g_total = EXCLUDED.g_total,
        actual_closed_date = EXCLUDED.actual_closed_date,
        checkpoints = EXCLUDED.checkpoints,
        repair_rows = EXCLUDED.repair_rows,
        part_rows = EXCLUDED.part_rows,
        created_by = EXCLUDED.created_by,
        created_by_email = EXCLUDED.created_by_email,
        updated_at = NOW();
    `;

    for (const card of cards) {
      if (!card.id) continue;
      await client.query(query, [
        card.id,
        card.jobNo || card.job_no || '',
        card.onlineJobCardNo || card.online_job_card_no || '',
        card.jobDate || card.job_date || '',
        card.dateTimeIn || card.date_time_in || '',
        card.dateTimeOut || card.date_time_out || '',
        card.expectedRepairTime || card.expected_repair_time || '',
        card.status || 'Open',
        card.custName || card.cust_name || '',
        card.fatherName || card.father_name || '',
        card.custAddr || card.cust_addr || '',
        card.village || '',
        card.mandal || '',
        card.ownerMob || card.owner_mob || '',
        card.driverMob || card.driver_mob || '',
        card.regdNo || card.regd_no || '',
        card.chassisNo || card.chassis_no || '',
        card.engineNo || card.engine_no || '',
        card.model || '',
        card.modelType || card.model_type || '',
        card.serialNo || card.serial_no || '',
        card.hourMeter || card.hour_meter || '',
        card.serviceType || card.service_type || '',
        card.freeServiceList || card.free_service_list || '',
        card.extraRepairs || card.extra_repairs || '',
        card.mechanic || '',
        card.wsIncharge || card.ws_incharge || '',
        card.serviceLocation || card.service_location || '',
        card.billNo || card.bill_no || '',
        card.reasonsForAnalysis || card.reasons_for_analysis || '',
        card.telecalling || '',
        card.warrantyOverride || card.warranty_override || 'auto',
        card.totalLabour || card.total_labour || '',
        card.warrantyMaterial || card.warranty_material || '',
        card.nonWarrantyMaterial || card.non_warranty_material || '',
        card.gTotal || card.g_total || '',
        card.actualClosedDate || card.actual_closed_date || '',
        JSON.stringify(card.checkpoints || []),
        JSON.stringify(card.repairRows || card.repair_rows || []),
        JSON.stringify(card.partRows || card.part_rows || []),
        card.createdBy || card.created_by || '',
        card.createdByEmail || card.created_by_email || '',
        card.createdAt || card.created_at || new Date().toISOString()
      ]);
    }

    await client.query('COMMIT');
    res.json({ success: true, count: cards.length });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error saving job cards bulk:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// =======================
// COMPLAINTS API
// =======================

// Get all complaints
app.get('/api/complaints', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM complaints ORDER BY created_at DESC');
    res.json({ success: true, data: result.rows });
  } catch (error: any) {
    console.error('Error fetching complaints:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Save/Update complaint
app.post('/api/complaints', async (req, res) => {
  const comp = req.body;
  if (!comp || !comp.id) {
    return res.status(400).json({ success: false, error: 'Complaint ID is required' });
  }

  try {
    const client = getPool();
    const query = `
      INSERT INTO complaints (
        id, complaint_no, date, customer_name, phone, village, mandal,
        tractor_model, chassis_no, hours, complaint_details, mechanic,
        status, job_card_no, closure_date, remarks, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, NOW())
      ON CONFLICT (id) DO UPDATE SET
        complaint_no = EXCLUDED.complaint_no,
        date = EXCLUDED.date,
        customer_name = EXCLUDED.customer_name,
        phone = EXCLUDED.phone,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        tractor_model = EXCLUDED.tractor_model,
        chassis_no = EXCLUDED.chassis_no,
        hours = EXCLUDED.hours,
        complaint_details = EXCLUDED.complaint_details,
        mechanic = EXCLUDED.mechanic,
        status = EXCLUDED.status,
        job_card_no = EXCLUDED.job_card_no,
        closure_date = EXCLUDED.closure_date,
        remarks = EXCLUDED.remarks,
        updated_at = NOW()
      RETURNING *;
    `;

    const result = await client.query(query, [
      comp.id,
      comp.complaintNo || comp.complaint_no || '',
      comp.date || '',
      comp.customerName || comp.customer_name || '',
      comp.phone || '',
      comp.village || '',
      comp.mandal || '',
      comp.tractorModel || comp.tractor_model || '',
      comp.chassisNo || comp.chassis_no || '',
      comp.hours || '',
      comp.complaintDetails || comp.complaint_details || '',
      comp.mechanic || '',
      comp.status || 'Pending',
      comp.jobCardNo || comp.job_card_no || '',
      comp.closureDate || comp.closure_date || '',
      comp.remarks || '',
      comp.createdAt || comp.created_at || new Date().toISOString()
    ]);

    res.json({ success: true, data: result.rows[0] });
  } catch (error: any) {
    console.error('Error saving complaint:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Delete complaint
app.delete('/api/complaints/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const client = getPool();
    await client.query('DELETE FROM complaints WHERE id = $1', [id]);
    res.json({ success: true });
  } catch (error: any) {
    console.error('Error deleting complaint:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Bulk upsert complaints
app.post('/api/complaints/bulk', async (req, res) => {
  const { complaints, replaceAll } = req.body;
  if (!Array.isArray(complaints)) {
    return res.status(400).json({ success: false, error: 'Complaints array required' });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    if (replaceAll) {
      await client.query('DELETE FROM complaints');
    }

    const query = `
      INSERT INTO complaints (
        id, complaint_no, date, customer_name, phone, village, mandal,
        tractor_model, chassis_no, hours, complaint_details, mechanic,
        status, job_card_no, closure_date, remarks, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, NOW())
      ON CONFLICT (id) DO UPDATE SET
        complaint_no = EXCLUDED.complaint_no,
        date = EXCLUDED.date,
        customer_name = EXCLUDED.customer_name,
        phone = EXCLUDED.phone,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        tractor_model = EXCLUDED.tractor_model,
        chassis_no = EXCLUDED.chassis_no,
        hours = EXCLUDED.hours,
        complaint_details = EXCLUDED.complaint_details,
        mechanic = EXCLUDED.mechanic,
        status = EXCLUDED.status,
        job_card_no = EXCLUDED.job_card_no,
        closure_date = EXCLUDED.closure_date,
        remarks = EXCLUDED.remarks,
        updated_at = NOW();
    `;

    for (const comp of complaints) {
      if (!comp.id) continue;
      await client.query(query, [
        comp.id,
        comp.complaintNo || comp.complaint_no || '',
        comp.date || '',
        comp.customerName || comp.customer_name || '',
        comp.phone || '',
        comp.village || '',
        comp.mandal || '',
        comp.tractorModel || comp.tractor_model || '',
        comp.chassisNo || comp.chassis_no || '',
        comp.hours || '',
        comp.complaintDetails || comp.complaint_details || '',
        comp.mechanic || '',
        comp.status || 'Pending',
        comp.jobCardNo || comp.job_card_no || '',
        comp.closureDate || comp.closure_date || '',
        comp.remarks || '',
        comp.createdAt || comp.created_at || new Date().toISOString()
      ]);
    }

    await client.query('COMMIT');
    res.json({ success: true, count: complaints.length });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error saving complaints bulk:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// =======================
// STAFF & ATTENDANCE API
// =======================

// Get all staff
app.get('/api/staff', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM staff ORDER BY name ASC');
    res.json({ success: true, data: result.rows });
  } catch (error: any) {
    console.error('Error fetching staff:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Save/Update staff
app.post('/api/staff', async (req, res) => {
  const st = req.body;
  if (!st || !st.id || !st.name || !st.role) {
    return res.status(400).json({ success: false, error: 'Staff ID, Name, and Role are required' });
  }

  try {
    const client = getPool();
    const query = `
      INSERT INTO staff (
        id, name, role, phone, active, assigned_supervisor, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, NOW())
      ON CONFLICT (id) DO UPDATE SET
        name = EXCLUDED.name,
        role = EXCLUDED.role,
        phone = EXCLUDED.phone,
        active = EXCLUDED.active,
        assigned_supervisor = EXCLUDED.assigned_supervisor,
        updated_at = NOW()
      RETURNING *;
    `;
    const result = await client.query(query, [
      st.id,
      st.name,
      st.role,
      st.phone || '',
      st.active !== undefined ? String(st.active) : 'true',
      st.assignedSupervisor || st.assigned_supervisor || '',
      st.createdAt || st.created_at || new Date().toISOString()
    ]);
    res.json({ success: true, data: result.rows[0] });
  } catch (error: any) {
    console.error('Error saving staff:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Delete staff
app.delete('/api/staff/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const client = getPool();
    await client.query('DELETE FROM staff WHERE id = $1', [id]);
    res.json({ success: true });
  } catch (error: any) {
    console.error('Error deleting staff:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Bulk staff
app.post('/api/staff/bulk', async (req, res) => {
  const { staff: staffList, replaceAll } = req.body;
  if (!Array.isArray(staffList)) {
    return res.status(400).json({ success: false, error: 'Staff array required' });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    if (replaceAll) {
      await client.query('DELETE FROM staff');
    }

    const query = `
      INSERT INTO staff (
        id, name, role, phone, active, assigned_supervisor, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, NOW())
      ON CONFLICT (id) DO UPDATE SET
        name = EXCLUDED.name,
        role = EXCLUDED.role,
        phone = EXCLUDED.phone,
        active = EXCLUDED.active,
        assigned_supervisor = EXCLUDED.assigned_supervisor,
        updated_at = NOW();
    `;

    for (const st of staffList) {
      if (!st.id || !st.name) continue;
      await client.query(query, [
        st.id,
        st.name,
        st.role || 'Mechanic',
        st.phone || '',
        st.active !== undefined ? String(st.active) : 'true',
        st.assignedSupervisor || st.assigned_supervisor || '',
        st.createdAt || st.created_at || new Date().toISOString()
      ]);
    }

    await client.query('COMMIT');
    res.json({ success: true, count: staffList.length });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error saving staff bulk:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// Attendance API
app.get('/api/attendance', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM staff_attendance');
    const recordsMap: Record<string, any> = {};
    result.rows.forEach(r => {
      recordsMap[r.date] = typeof r.records === 'string' ? JSON.parse(r.records) : r.records;
    });
    res.json({ success: true, data: recordsMap });
  } catch (error: any) {
    console.error('Error fetching attendance:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/attendance', async (req, res) => {
  const { date, records } = req.body;
  if (!date || !records) {
    return res.status(400).json({ success: false, error: 'Date and records required' });
  }

  try {
    const client = getPool();
    const query = `
      INSERT INTO staff_attendance (date, records, updated_at)
      VALUES ($1, $2, NOW())
      ON CONFLICT (date) DO UPDATE SET
        records = EXCLUDED.records,
        updated_at = NOW();
    `;
    await client.query(query, [date, JSON.stringify(records)]);
    res.json({ success: true });
  } catch (error: any) {
    console.error('Error saving attendance:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// App settings API
app.get('/api/settings', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM app_settings');
    const settingsMap: Record<string, string> = {};
    result.rows.forEach(r => {
      settingsMap[r.key] = r.value;
    });
    res.json({ success: true, data: settingsMap });
  } catch (error: any) {
    console.error('Error fetching settings:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/settings', async (req, res) => {
  const { key, value } = req.body;
  if (!key) return res.status(400).json({ success: false, error: 'Key required' });

  try {
    const client = getPool();
    const query = `
      INSERT INTO app_settings (key, value, updated_at)
      VALUES ($1, $2, NOW())
      ON CONFLICT (key) DO UPDATE SET
        value = EXCLUDED.value,
        updated_at = NOW();
    `;
    await client.query(query, [key, String(value)]);
    res.json({ success: true });
  } catch (error: any) {
    console.error('Error saving setting:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Full Master Backup JSON/Export endpoint
app.get('/api/master-backup', async (req, res) => {
  try {
    const client = getPool();
    const [custRes, sparesRes, jcRes, compRes, staffRes, attRes, setRes] = await Promise.all([
      client.query('SELECT * FROM customers'),
      client.query('SELECT * FROM spares'),
      client.query('SELECT * FROM jobcards'),
      client.query('SELECT * FROM complaints'),
      client.query('SELECT * FROM staff'),
      client.query('SELECT * FROM staff_attendance'),
      client.query('SELECT * FROM app_settings'),
    ]);

    res.json({
      success: true,
      customers: custRes.rows,
      spares: sparesRes.rows,
      jobcards: jcRes.rows,
      complaints: compRes.rows,
      staff: staffRes.rows,
      attendance: attRes.rows,
      settings: setRes.rows
    });
  } catch (error: any) {
    console.error('Error generating master backup:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Eicher standalone server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
