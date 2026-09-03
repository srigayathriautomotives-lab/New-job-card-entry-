import express from 'express';
import cors from 'cors';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import { pool } from './src/db/index.ts';

const app = express();
const PORT = Number(process.env.PORT || 3000);

app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

function getPool() {
  return pool;
}

async function ensureDbTables() {
  try {
    const client = getPool();
    await client.query(`
      CREATE TABLE IF NOT EXISTS jobcards (
        id TEXT PRIMARY KEY,
        job_no TEXT,
        online_job_card_no TEXT,
        job_date TEXT,
        date_time_in TEXT,
        date_time_out TEXT,
        expected_repair_time TEXT,
        status TEXT,
        cust_name TEXT,
        father_name TEXT,
        cust_addr TEXT,
        village TEXT,
        mandal TEXT,
        owner_mob TEXT,
        driver_mob TEXT,
        regd_no TEXT,
        chassis_no TEXT,
        engine_no TEXT,
        model TEXT,
        model_type TEXT,
        serial_no TEXT,
        hour_meter TEXT,
        service_type TEXT,
        free_service_list TEXT,
        extra_repairs TEXT,
        mechanic TEXT,
        ws_incharge TEXT,
        service_location TEXT,
        bill_no TEXT,
        reasons_for_analysis TEXT,
        telecalling TEXT,
        warranty_override TEXT,
        total_labour TEXT,
        warranty_material TEXT,
        non_warranty_material TEXT,
        g_total TEXT,
        actual_closed_date TEXT,
        branch TEXT,
        history_file_no TEXT,
        complaint_date TEXT,
        install_date TEXT,
        date_of_delivery TEXT,
        dist_dealership TEXT,
        full_data TEXT,
        checkpoints TEXT,
        repair_rows TEXT,
        part_rows TEXT,
        created_by TEXT,
        created_by_email TEXT,
        created_at TEXT,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);

    await client.query(`
      CREATE TABLE IF NOT EXISTS service_camps (
        id TEXT PRIMARY KEY,
        dealership_code TEXT,
        branch TEXT,
        mandal TEXT,
        village TEXT,
        camp_date TEXT,
        target_tractors TEXT,
        supervisor TEXT,
        mechanic TEXT,
        status TEXT,
        service_type_expected TEXT,
        offers TEXT,
        contact_person TEXT,
        contact_phone TEXT,
        notes TEXT,
        attended_count TEXT,
        created_at TEXT,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Ensure staff table and columns exist
    try {
      await client.query(`
        CREATE TABLE IF NOT EXISTS staff (
          id TEXT PRIMARY KEY,
          name TEXT NOT NULL,
          role TEXT NOT NULL,
          phone TEXT,
          father_name TEXT,
          village TEXT,
          mandal TEXT,
          mobile_number TEXT,
          date_of_joining TEXT,
          supervisor TEXT,
          active TEXT DEFAULT 'true',
          assigned_supervisor TEXT,
          created_at TEXT,
          updated_at TIMESTAMP DEFAULT NOW()
        );
      `);
    } catch (e) {
      console.warn('Notice ensuring staff table:', e);
    }

    // Ensure customers table
    await client.query(`
      CREATE TABLE IF NOT EXISTS customers (
        id SERIAL PRIMARY KEY,
        chassis_key TEXT NOT NULL UNIQUE,
        chassis_no TEXT NOT NULL,
        cust_name TEXT,
        father_name TEXT,
        cust_addr TEXT,
        village TEXT,
        mandal TEXT,
        owner_mob TEXT,
        driver_mob TEXT,
        regd_no TEXT,
        engine_no TEXT,
        tractor_model TEXT,
        date_of_delivery TEXT,
        followup_history TEXT,
        full_data TEXT,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Ensure spares table
    await client.query(`
      CREATE TABLE IF NOT EXISTS spares (
        id SERIAL PRIMARY KEY,
        part_key TEXT NOT NULL UNIQUE,
        part_no TEXT NOT NULL,
        part_desc TEXT,
        mrp TEXT,
        category TEXT,
        full_data TEXT,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Ensure complaints table
    await client.query(`
      CREATE TABLE IF NOT EXISTS complaints (
        id TEXT PRIMARY KEY,
        complaint_no TEXT,
        date TEXT,
        customer_name TEXT,
        phone TEXT,
        village TEXT,
        mandal TEXT,
        tractor_model TEXT,
        chassis_no TEXT,
        hours TEXT,
        complaint_details TEXT,
        mechanic TEXT,
        status TEXT,
        job_card_no TEXT,
        closure_date TEXT,
        remarks TEXT,
        created_at TEXT,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Ensure staff_attendance table
    await client.query(`
      CREATE TABLE IF NOT EXISTS staff_attendance (
        id SERIAL PRIMARY KEY,
        date TEXT NOT NULL UNIQUE,
        records TEXT NOT NULL,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Ensure app_settings table
    await client.query(`
      CREATE TABLE IF NOT EXISTS app_settings (
        id SERIAL PRIMARY KEY,
        key TEXT NOT NULL UNIQUE,
        value TEXT NOT NULL,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Ensure columns exist on jobcards table
    await client.query(`
      ALTER TABLE jobcards ADD COLUMN IF NOT EXISTS branch TEXT;
      ALTER TABLE jobcards ADD COLUMN IF NOT EXISTS history_file_no TEXT;
      ALTER TABLE jobcards ADD COLUMN IF NOT EXISTS complaint_date TEXT;
      ALTER TABLE jobcards ADD COLUMN IF NOT EXISTS install_date TEXT;
      ALTER TABLE jobcards ADD COLUMN IF NOT EXISTS date_of_delivery TEXT;
      ALTER TABLE jobcards ADD COLUMN IF NOT EXISTS dist_dealership TEXT;
      ALTER TABLE jobcards ADD COLUMN IF NOT EXISTS full_data TEXT;
    `);

    console.log('Database tables verified/created successfully.');
  } catch (err: any) {
    if (err && err.message && err.message.includes('permission denied for schema public')) {
      console.log('Notice: Database user does not have DDL permission on schema public; assuming tables are pre-configured.');
    } else {
      console.warn('DB initialization check:', err);
    }
  }
}
ensureDbTables();

// Clear database API
app.post('/api/database/clear', async (req, res) => {
  const client = await getPool().connect();
  console.log('Database clearing initiated on configured PostgreSQL database');
  try {
    await client.query('BEGIN');
    await client.query('SET session_replication_role = replica;');
    await client.query('TRUNCATE TABLE customers, jobcards, complaints, spares, staff, staff_attendance, app_settings, service_camps RESTART IDENTITY CASCADE');
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

// Intelligent fuzzy field normalizer for Excel / CSV / JSON imports
function normKey(str: any): string {
  return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function extractField(obj: any, candidates: string[]): string {
  if (!obj || typeof obj !== 'object') return '';
  const normalizedObj: Record<string, any> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (v != null && String(v).trim() !== '') {
      normalizedObj[normKey(k)] = String(v).trim();
    }
  }

  // 1. Exact normalized candidate match
  for (const c of candidates) {
    const nc = normKey(c);
    if (normalizedObj[nc] !== undefined && normalizedObj[nc] !== '') {
      return normalizedObj[nc];
    }
  }

  // 2. Substring matching in keys
  for (const [k, v] of Object.entries(normalizedObj)) {
    for (const c of candidates) {
      const nc = normKey(c);
      if (nc.length >= 3 && (k.includes(nc) || nc.includes(k)) && v !== '') {
        return v;
      }
    }
  }

  return '';
}

// Extract customer object from any Excel / raw row format
function normalizeCustomerRow(r: any, index: number, seenKeys: Set<string>) {
  const chassisNo = extractField(r, [
    'chassisNo', 'chassis_no', 'chassis', 'chassisnumber', 'chassisnum', 'chassiscode',
    'vin', 'vinnumber', 'frameno', 'framenumber', 'tractorslno', 'tractorchassis',
    'serialno', 'serialnumber', 'slno', 'sno', 'trchassis'
  ]) || (r.chassisNo ? String(r.chassisNo).trim() : '');

  const custName = extractField(r, [
    'custName', 'cust_name', 'customername', 'nameofcustomer', 'name', 'ownername',
    'farmername', 'clientname', 'partyname', 'customer', 'cust'
  ]) || (r.custName ? String(r.custName).trim() : '');

  const fatherName = extractField(r, [
    'fatherName', 'father_name', 'father', 'fathersname', 'careof', 'co', 'so', 'wo', 'do',
    'husbandname', 'parentname', 'guardian'
  ]) || (r.fatherName ? String(r.fatherName).trim() : '');

  const custAddr = extractField(r, [
    'custAddr', 'cust_addr', 'address', 'customeraddress', 'fulladdress', 'location',
    'residence', 'place', 'addr'
  ]) || (r.custAddr ? String(r.custAddr).trim() : '');

  const village = extractField(r, [
    'village', 'vill', 'town', 'city', 'habitation', 'gramam', 'ooru'
  ]) || (r.village ? String(r.village).trim() : '');

  const mandal = extractField(r, [
    'mandal', 'mandalname', 'taluk', 'tehsil', 'block', 'district', 'dist'
  ]) || (r.mandal ? String(r.mandal).trim() : '');

  const ownerMob = extractField(r, [
    'ownerMob', 'owner_mob', 'mobile', 'mobilenumber', 'phone', 'phonenumber', 'contact',
    'contactno', 'custphone', 'cell', 'cellno', 'phno', 'tel', 'phone1'
  ]) || (r.ownerMob ? String(r.ownerMob).trim() : '');

  const driverMob = extractField(r, [
    'driverMob', 'driver_mob', 'driverphone', 'drivernumber', 'alternatemobile', 'altphone',
    'altmobile', 'secondmobile', 'phone2', 'driver'
  ]) || (r.driverMob ? String(r.driverMob).trim() : '');

  const regdNo = extractField(r, [
    'regdNo', 'regd_no', 'registrationno', 'regnumber', 'vehicleno', 'tractorregdno',
    'regno', 'rcno', 'plateno'
  ]) || (r.regdNo ? String(r.regdNo).trim() : '');

  const engineNo = extractField(r, [
    'engineNo', 'engine_no', 'enginenumber', 'engno', 'motorno'
  ]) || (r.engineNo ? String(r.engineNo).trim() : '');

  const tractorModel = extractField(r, [
    'tractorModel', 'tractor_model', 'model', 'modeltype', 'modelname', 'variant',
    'hp', 'horse_power', 'make', 'tractortype'
  ]) || (r.tractorModel ? String(r.tractorModel).trim() : '');

  const dateOfDelivery = extractField(r, [
    'dateOfDelivery', 'date_of_delivery', 'deliverydate', 'installdate', 'doi', 'dop',
    'purchasedate', 'saledate', 'invoicedate', 'billdate', 'dod', 'delivery_date', 'date'
  ]) || (r.dateOfDelivery ? String(r.dateOfDelivery).trim() : '');

  // Calculate unique chassisKey so NO row is ever skipped
  let baseKey = normKey(chassisNo);
  if (!baseKey) {
    if (ownerMob) baseKey = `mob_${normKey(ownerMob)}`;
    else if (regdNo) baseKey = `reg_${normKey(regdNo)}`;
    else if (custName) baseKey = `name_${normKey(custName)}`;
    else baseKey = `cust_row_${index + 1}`;
  }

  let finalKey = baseKey;
  if (seenKeys) {
    // Only append duplicate suffix if there is NO real chassis number provided
    if (!normKey(chassisNo)) {
      let counter = 1;
      while (seenKeys.has(finalKey)) {
        counter++;
        finalKey = `${baseKey}_dup${counter}`;
      }
    }
    seenKeys.add(finalKey);
  }

  let followupHistory = [];
  try {
    followupHistory = typeof r.followupHistory === 'string'
      ? JSON.parse(r.followupHistory)
      : (Array.isArray(r.followupHistory) ? r.followupHistory : (Array.isArray(r.followup_history) ? r.followup_history : []));
  } catch {}

  const fullData = r.full_data || r.fullData || {
    ...r,
    chassisNo,
    custName,
    fatherName,
    custAddr: custAddr || (village ? (mandal ? `${village}, ${mandal}` : village) : ''),
    village,
    mandal,
    ownerMob,
    driverMob,
    regdNo,
    engineNo,
    tractorModel,
    dateOfDelivery,
    __chassisDisplay: chassisNo,
    __custNameDisplay: custName,
    __custPhoneDisplay: ownerMob,
    __custAddrDisplay: custAddr || village
  };

  return {
    chassisKey: finalKey,
    chassisNo: chassisNo || finalKey,
    custName,
    fatherName,
    custAddr: custAddr || (village ? (mandal ? `${village}, ${mandal}` : village) : ''),
    village,
    mandal,
    ownerMob,
    driverMob,
    regdNo,
    engineNo,
    tractorModel,
    dateOfDelivery,
    followupHistory: JSON.stringify(followupHistory),
    fullData: JSON.stringify(fullData)
  };
}

// Extract spare row from any Excel / raw format
function normalizeSpareRow(r: any, index: number, seenKeys: Set<string>) {
  const partNo = extractField(r, [
    'partNo', 'part_no', 'partnumber', 'itemcode', 'sparepartno', 'spareno',
    'partcode', 'materialno', 'itemnumber', 'part', 'code'
  ]) || (r.partNo ? String(r.partNo).trim() : '');

  const partDesc = extractField(r, [
    'partDesc', 'part_desc', 'desc', 'description', 'partname', 'itemdesc',
    'itemname', 'materialdescription', 'sparedesc', 'details', 'name'
  ]) || (r.partDesc ? String(r.partDesc).trim() : '');

  const mrp = extractField(r, [
    'mrp', 'rate', 'price', 'amount', 'unitprice', 'cost', 'val', 'standardrate'
  ]) || (r.mrp != null ? String(r.mrp).trim() : '');

  const category = extractField(r, [
    'category', 'group', 'type', 'partcategory', 'section', 'division'
  ]) || (r.category ? String(r.category).trim() : '');

  let baseKey = normKey(partNo);
  if (!baseKey) {
    if (partDesc) baseKey = `desc_${normKey(partDesc).substring(0, 20)}`;
    else baseKey = `part_row_${index + 1}`;
  }

  let finalKey = baseKey;
  let counter = 1;
  while (seenKeys.has(finalKey)) {
    counter++;
    finalKey = `${baseKey}_dup${counter}`;
  }
  seenKeys.add(finalKey);

  const fullData = r.full_data || r.fullData || {
    ...r,
    partNo,
    partDesc,
    mrp,
    category,
    __partNoDisplay: partNo
  };

  return {
    partKey: finalKey,
    partNo: partNo || finalKey,
    partDesc,
    mrp,
    category,
    fullData: JSON.stringify(fullData)
  };
}

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

  if (rows.length === 0) {
    return res.json({ success: true, count: 0 });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');

    if (replaceAll) {
      await client.query('DELETE FROM customers');
    }

    const seenKeys = new Set<string>();
    const rowMap = new Map<string, any>();
    rows.forEach((r, i) => {
      const norm = normalizeCustomerRow(r, i, seenKeys);
      if (rowMap.has(norm.chassisKey)) {
        const existing = rowMap.get(norm.chassisKey);
        rowMap.set(norm.chassisKey, { ...existing, ...norm });
      } else {
        rowMap.set(norm.chassisKey, norm);
      }
    });
    const normalizedRows = Array.from(rowMap.values());

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

    // Process in parallel batches of 50 for rapid upload speed
    const batchSize = 50;
    for (let i = 0; i < normalizedRows.length; i += batchSize) {
      const chunk = normalizedRows.slice(i, i + batchSize);
      await Promise.all(
        chunk.map((item) =>
          client.query(insertQuery, [
            item.chassisKey,
            item.chassisNo,
            item.custName,
            item.fatherName,
            item.custAddr,
            item.village,
            item.mandal,
            item.ownerMob,
            item.driverMob,
            item.regdNo,
            item.engineNo,
            item.tractorModel,
            item.dateOfDelivery,
            item.followupHistory,
            item.fullData
          ])
        )
      );
    }

    await client.query('COMMIT');
    res.json({ success: true, count: normalizedRows.length });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error saving customers bulk:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// Delete all customers
app.post('/api/database/delete-all-customers', async (req, res) => {
  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM customers');
    await client.query('COMMIT');
    res.json({ success: true, message: 'All customers deleted successfully' });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error deleting all customers:', error);
    res.status(500).json({ success: false, error: 'Database error' });
  } finally {
    client.release();
  }
});

// Upsert single customer
app.post('/api/customers', async (req, res) => {
  const r = req.body;
  const seenKeys = new Set<string>();
  const item = normalizeCustomerRow(r, 0, seenKeys);

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
      item.chassisKey,
      item.chassisNo,
      item.custName,
      item.fatherName,
      item.custAddr,
      item.village,
      item.mandal,
      item.ownerMob,
      item.driverMob,
      item.regdNo,
      item.engineNo,
      item.tractorModel,
      item.dateOfDelivery,
      item.followupHistory,
      item.fullData
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

  if (rows.length === 0 && !replaceAll) {
    return res.json({ success: true, count: 0 });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');

    if (replaceAll) {
      await client.query('DELETE FROM spares');
    }

    const seenKeys = new Set<string>();
    let normalizedRows = rows.map((r, i) => normalizeSpareRow(r, i, seenKeys));

    // If merging (not replacing), we should deduplicate by part_no if possible
    if (!replaceAll) {
      const existing = await client.query('SELECT part_no, part_key FROM spares');
      const existingPartNos = new Map(existing.rows.map(r => [r.part_no, r.part_key]));
      
      normalizedRows = normalizedRows.map(row => {
        if (existingPartNos.has(row.partNo)) {
          return { ...row, partKey: existingPartNos.get(row.partNo) };
        }
        return row;
      });
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

    const batchSize = 50;
    for (let i = 0; i < normalizedRows.length; i += batchSize) {
      const chunk = normalizedRows.slice(i, i + batchSize);
      await Promise.all(
        chunk.map((item) =>
          client.query(insertQuery, [
            item.partKey,
            item.partNo,
            item.partDesc,
            item.mrp,
            item.category,
            item.fullData
          ])
        )
      );
    }

    await client.query('COMMIT');
    res.json({ success: true, count: normalizedRows.length });
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
        branch, history_file_no, complaint_date, install_date, date_of_delivery, dist_dealership, full_data,
        checkpoints, repair_rows, part_rows, created_by, created_by_email, created_at, updated_at
      ) VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16,
        $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, $28, $29, $30,
        $31, $32, $33, $34, $35, $36, $37, $38, $39, $40, $41, $42, $43, $44,
        $45, $46, $47, $48, $49, $50, NOW()
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
        branch = EXCLUDED.branch,
        history_file_no = EXCLUDED.history_file_no,
        complaint_date = EXCLUDED.complaint_date,
        install_date = EXCLUDED.install_date,
        date_of_delivery = EXCLUDED.date_of_delivery,
        dist_dealership = EXCLUDED.dist_dealership,
        full_data = EXCLUDED.full_data,
        checkpoints = EXCLUDED.checkpoints,
        repair_rows = EXCLUDED.repair_rows,
        part_rows = EXCLUDED.part_rows,
        created_by = EXCLUDED.created_by,
        created_by_email = EXCLUDED.created_by_email,
        updated_at = NOW()
      RETURNING *;
    `;

    const fullData = card.fullData || card.full_data || card;

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
      card.mechanic || card.technicianName || '',
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
      card.branch || '',
      card.historyFileNo || card.history_file_no || '',
      card.complaintDate || card.complaint_date || '',
      card.installDate || card.install_date || card.dateOfDelivery || card.date_of_delivery || '',
      card.dateOfDelivery || card.date_of_delivery || card.installDate || card.install_date || '',
      card.distDealership || card.dist_dealership || '',
      typeof fullData === 'object' ? JSON.stringify(fullData) : String(fullData || ''),
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
        branch, history_file_no, complaint_date, install_date, date_of_delivery, dist_dealership, full_data,
        checkpoints, repair_rows, part_rows, created_by, created_by_email, created_at, updated_at
      ) VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16,
        $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, $28, $29, $30,
        $31, $32, $33, $34, $35, $36, $37, $38, $39, $40, $41, $42, $43, $44,
        $45, $46, $47, $48, $49, $50, NOW()
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
        branch = EXCLUDED.branch,
        history_file_no = EXCLUDED.history_file_no,
        complaint_date = EXCLUDED.complaint_date,
        install_date = EXCLUDED.install_date,
        date_of_delivery = EXCLUDED.date_of_delivery,
        dist_dealership = EXCLUDED.dist_dealership,
        full_data = EXCLUDED.full_data,
        checkpoints = EXCLUDED.checkpoints,
        repair_rows = EXCLUDED.repair_rows,
        part_rows = EXCLUDED.part_rows,
        created_by = EXCLUDED.created_by,
        created_by_email = EXCLUDED.created_by_email,
        updated_at = NOW();
    `;

    for (const card of cards) {
      if (!card.id) continue;
      const fullData = card.fullData || card.full_data || card;
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
        card.mechanic || card.technicianName || '',
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
        card.branch || '',
        card.historyFileNo || card.history_file_no || '',
        card.complaintDate || card.complaint_date || '',
        card.installDate || card.install_date || card.dateOfDelivery || card.date_of_delivery || '',
        card.dateOfDelivery || card.date_of_delivery || card.installDate || card.install_date || '',
        card.distDealership || card.dist_dealership || '',
        typeof fullData === 'object' ? JSON.stringify(fullData) : String(fullData || ''),
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
    const mapped = result.rows.map(r => ({
      id: r.id,
      name: r.name,
      role: r.role,
      phone: r.phone || r.mobile_number || '',
      mobileNumber: r.mobile_number || r.phone || '',
      fatherName: r.father_name || '',
      village: r.village || '',
      mandal: r.mandal || '',
      dateOfJoining: r.date_of_joining || '',
      supervisor: r.supervisor || r.assigned_supervisor || '',
      assignedSupervisor: r.assigned_supervisor || r.supervisor || '',
      active: r.active === 'true' || r.active === true || r.active === undefined || r.active === '1',
      createdAt: r.created_at || '',
      updatedAt: r.updated_at || ''
    }));
    res.json({ success: true, data: mapped });
  } catch (error: any) {
    console.error('Error fetching staff:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Save/Update staff
app.post('/api/staff', async (req, res) => {
  const st = req.body;
  if (!st || !st.id || !st.name) {
    return res.status(400).json({ success: false, error: 'Staff ID and Name are required' });
  }

  try {
    const client = getPool();
    const query = `
      INSERT INTO staff (
        id, name, role, phone, father_name, village, mandal, mobile_number, date_of_joining, supervisor, active, assigned_supervisor, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, NOW())
      ON CONFLICT (id) DO UPDATE SET
        name = EXCLUDED.name,
        role = EXCLUDED.role,
        phone = EXCLUDED.phone,
        father_name = EXCLUDED.father_name,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        mobile_number = EXCLUDED.mobile_number,
        date_of_joining = EXCLUDED.date_of_joining,
        supervisor = EXCLUDED.supervisor,
        active = EXCLUDED.active,
        assigned_supervisor = EXCLUDED.assigned_supervisor,
        updated_at = NOW()
      RETURNING *;
    `;
    const result = await client.query(query, [
      st.id,
      st.name,
      st.role || 'mechanic',
      st.phone || st.mobileNumber || '',
      st.fatherName || st.father_name || '',
      st.village || '',
      st.mandal || '',
      st.mobileNumber || st.mobile_number || st.phone || '',
      st.dateOfJoining || st.date_of_joining || '',
      st.supervisor || st.assignedSupervisor || '',
      st.active !== undefined ? String(st.active) : 'true',
      st.assignedSupervisor || st.assigned_supervisor || st.supervisor || '',
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
  const staffList = Array.isArray(req.body.staffList) ? req.body.staffList : (Array.isArray(req.body.staff) ? req.body.staff : (Array.isArray(req.body) ? req.body : []));
  const replaceAll = req.body.replaceAll;
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
        id, name, role, phone, father_name, village, mandal, mobile_number, date_of_joining, supervisor, active, assigned_supervisor, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, NOW())
      ON CONFLICT (id) DO UPDATE SET
        name = EXCLUDED.name,
        role = EXCLUDED.role,
        phone = EXCLUDED.phone,
        father_name = EXCLUDED.father_name,
        village = EXCLUDED.village,
        mandal = EXCLUDED.mandal,
        mobile_number = EXCLUDED.mobile_number,
        date_of_joining = EXCLUDED.date_of_joining,
        supervisor = EXCLUDED.supervisor,
        active = EXCLUDED.active,
        assigned_supervisor = EXCLUDED.assigned_supervisor,
        updated_at = NOW();
    `;

    for (const st of staffList) {
      if (!st.id || !st.name) continue;
      await client.query(query, [
        st.id,
        st.name,
        st.role || 'mechanic',
        st.phone || st.mobileNumber || '',
        st.fatherName || st.father_name || '',
        st.village || '',
        st.mandal || '',
        st.mobileNumber || st.mobile_number || st.phone || '',
        st.dateOfJoining || st.date_of_joining || '',
        st.supervisor || st.assignedSupervisor || '',
        st.active !== undefined ? String(st.active) : 'true',
        st.assignedSupervisor || st.assigned_supervisor || st.supervisor || '',
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

// Service Camps API
app.get('/api/service-camps', async (req, res) => {
  try {
    const client = getPool();
    const result = await client.query('SELECT * FROM service_camps ORDER BY camp_date ASC, created_at DESC');
    res.json({ success: true, data: result.rows });
  } catch (error: any) {
    console.error('Error fetching service camps:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/service-camps', async (req, res) => {
  const camp = req.body;
  if (!camp || !camp.id) {
    return res.status(400).json({ success: false, error: 'Camp ID is required' });
  }

  try {
    const client = getPool();
    const query = `
      INSERT INTO service_camps (
        id, dealership_code, branch, mandal, village, camp_date,
        target_tractors, supervisor, mechanic, status, service_type_expected,
        offers, contact_person, contact_phone, notes, attended_count, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, NOW())
      ON CONFLICT (id) DO UPDATE SET
        dealership_code = EXCLUDED.dealership_code,
        branch = EXCLUDED.branch,
        mandal = EXCLUDED.mandal,
        village = EXCLUDED.village,
        camp_date = EXCLUDED.camp_date,
        target_tractors = EXCLUDED.target_tractors,
        supervisor = EXCLUDED.supervisor,
        mechanic = EXCLUDED.mechanic,
        status = EXCLUDED.status,
        service_type_expected = EXCLUDED.service_type_expected,
        offers = EXCLUDED.offers,
        contact_person = EXCLUDED.contact_person,
        contact_phone = EXCLUDED.contact_phone,
        notes = EXCLUDED.notes,
        attended_count = EXCLUDED.attended_count,
        updated_at = NOW()
      RETURNING *;
    `;
    const result = await client.query(query, [
      camp.id,
      camp.dealershipCode || camp.dealership_code || '',
      camp.branch || '',
      camp.mandal || '',
      camp.village || '',
      camp.campDate || camp.camp_date || '',
      camp.targetTractors || camp.target_tractors || '',
      camp.supervisor || '',
      camp.mechanic || '',
      camp.status || 'Upcoming',
      camp.serviceTypeExpected || camp.service_type_expected || '',
      camp.offers || '',
      camp.contactPerson || camp.contact_person || '',
      camp.contactPhone || camp.contact_phone || '',
      camp.notes || '',
      camp.attendedCount || camp.attended_count || '',
      camp.createdAt || camp.created_at || new Date().toISOString()
    ]);
    res.json({ success: true, data: result.rows[0] });
  } catch (error: any) {
    console.error('Error saving service camp:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.delete('/api/service-camps/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const client = getPool();
    await client.query('DELETE FROM service_camps WHERE id = $1', [id]);
    res.json({ success: true });
  } catch (error: any) {
    console.error('Error deleting service camp:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/service-camps/bulk', async (req, res) => {
  const { camps, replaceAll } = req.body;
  if (!Array.isArray(camps)) {
    return res.status(400).json({ success: false, error: 'Camps array required' });
  }

  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    if (replaceAll) {
      await client.query('DELETE FROM service_camps');
    }

    const query = `
      INSERT INTO service_camps (
        id, dealership_code, branch, mandal, village, camp_date,
        target_tractors, supervisor, mechanic, status, service_type_expected,
        offers, contact_person, contact_phone, notes, attended_count, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, NOW())
      ON CONFLICT (id) DO UPDATE SET
        dealership_code = EXCLUDED.dealership_code,
        branch = EXCLUDED.branch,
        mandal = EXCLUDED.mandal,
        village = EXCLUDED.village,
        camp_date = EXCLUDED.camp_date,
        target_tractors = EXCLUDED.target_tractors,
        supervisor = EXCLUDED.supervisor,
        mechanic = EXCLUDED.mechanic,
        status = EXCLUDED.status,
        service_type_expected = EXCLUDED.service_type_expected,
        offers = EXCLUDED.offers,
        contact_person = EXCLUDED.contact_person,
        contact_phone = EXCLUDED.contact_phone,
        notes = EXCLUDED.notes,
        attended_count = EXCLUDED.attended_count,
        updated_at = NOW();
    `;

    for (const camp of camps) {
      if (!camp.id) continue;
      await client.query(query, [
        camp.id,
        camp.dealershipCode || camp.dealership_code || '',
        camp.branch || '',
        camp.mandal || '',
        camp.village || '',
        camp.campDate || camp.camp_date || '',
        camp.targetTractors || camp.target_tractors || '',
        camp.supervisor || '',
        camp.mechanic || '',
        camp.status || 'Upcoming',
        camp.serviceTypeExpected || camp.service_type_expected || '',
        camp.offers || '',
        camp.contactPerson || camp.contact_person || '',
        camp.contactPhone || camp.contact_phone || '',
        camp.notes || '',
        camp.attendedCount || camp.attended_count || '',
        camp.createdAt || camp.created_at || new Date().toISOString()
      ]);
    }

    await client.query('COMMIT');
    res.json({ success: true, count: camps.length });
  } catch (error: any) {
    await client.query('ROLLBACK');
    console.error('Error saving service camps bulk:', error);
    res.status(500).json({ success: false, error: error.message });
  } finally {
    client.release();
  }
});

// Full Master Backup JSON/Export endpoint
app.get('/api/master-backup', async (req, res) => {
  try {
    const client = getPool();
    const [custRes, sparesRes, jcRes, compRes, staffRes, attRes, setRes, campRes] = await Promise.all([
      client.query('SELECT * FROM customers'),
      client.query('SELECT * FROM spares'),
      client.query('SELECT * FROM jobcards'),
      client.query('SELECT * FROM complaints'),
      client.query('SELECT * FROM staff'),
      client.query('SELECT * FROM staff_attendance'),
      client.query('SELECT * FROM app_settings'),
      client.query('SELECT * FROM service_camps'),
    ]);

    res.json({
      success: true,
      customers: custRes.rows,
      spares: sparesRes.rows,
      jobcards: jcRes.rows,
      complaints: compRes.rows,
      staff: staffRes.rows,
      attendance: attRes.rows,
      settings: setRes.rows,
      serviceCamps: campRes.rows
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
