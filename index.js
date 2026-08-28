import express from 'express';
import cors from 'cors';
import { createClient } from '@supabase/supabase-js';
import multer from 'multer';
import * as XLSX from 'xlsx';

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json());

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

app.get('/api/v1/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.get('/api/v1/trips', async (req, res) => {
  try {
    const { search, sortBy = 'trip_date', sortOrder = 'desc' } = req.query;
    let query = supabase.from('trip_records').select('*');
    
    if (search) {
      query = query.or(`passenger.ilike.%${search}%,destination.ilike.%${search}%,driver.ilike.%${search}%`);
    }
    
    const order = sortOrder === 'asc' ? 'asc' : 'desc';
    query = query.order(sortBy, { ascending: order === 'asc' });
    
    const { data, error } = await query;
    if (error) throw error;
    
    res.json({ success: true, data: data || [] });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/v1/trips', async (req, res) => {
  try {
    const { tripDate, tripTime, passenger, destination, driver, tripType, mileage, amount, remark } = req.body;
    const tripDateTime = tripTime ? `${tripDate}T${tripTime}:00` : `${tripDate}T00:00:00`;
    
    const { data, error } = await supabase
      .from('trip_records')
      .insert([{
        trip_date: tripDateTime,
        passenger,
        destination,
        driver: driver || '待定',
        trip_type: tripType || '公务用车',
        mileage: mileage || 0,
        amount: amount || 0,
        remark: remark || ''
      }])
      .select()
      .single();
    
    if (error) throw error;
    res.json({ success: true, data });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.put('/api/v1/trips/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { tripDate, tripTime, passenger, destination, driver, tripType, mileage, amount, remark } = req.body;
    const tripDateTime = tripTime ? `${tripDate}T${tripTime}:00` : `${tripDate}T00:00:00`;
    
    const { data, error } = await supabase
      .from('trip_records')
      .update({
        trip_date: tripDateTime,
        passenger,
        destination,
        driver: driver || '待定',
        trip_type: tripType || '公务用车',
        mileage: mileage || 0,
        amount: amount || 0,
        remark: remark || '',
        updated_at: new Date().toISOString()
      })
      .eq('id', id)
      .select()
      .single();
    
    if (error) throw error;
    res.json({ success: true, data });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.delete('/api/v1/trips/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { data, error } = await supabase
      .from('trip_records')
      .delete()
      .eq('id', id)
      .select()
      .single();
    
    if (error) throw error;
    res.json({ success: true, data });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/v1/trips/export', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('trip_records')
      .select('*')
      .order('trip_date', { ascending: false });
    
    if (error) throw error;
    
    const excelData = (data || []).map(record => ({
      '出车日期时间': formatDateTime(record.trip_date),
      '用车人': record.passenger,
      '出车地点事由': record.destination,
      '驾驶员': record.driver,
      '出车方式': record.trip_type,
      '公里数': record.mileage,
      '金额': record.amount,
      '备注': record.remark || ''
    }));
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData);
    XLSX.utils.book_append_sheet(wb, ws, '出车记录');
    
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=trips.xlsx');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/v1/trips/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: '请上传文件' });
    }
    
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    
    const results = { total: jsonData.length, success: 0, failed: 0, errors: [] };
    
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      const rowNum = i + 2;
      
      try {
        const tripDate = row['出车日期时间'];
        const passenger = row['用车人'];
        const destination = row['出车地点事由'];
        
        if (!tripDate || !passenger || !destination) {
          results.failed++;
          results.errors.push(`第${rowNum}行：缺少必填字段`);
          continue;
        }
        
        const { error } = await supabase
          .from('trip_records')
          .insert([{
            trip_date: tripDate,
            passenger,
            destination,
            driver: row['驾驶员'] || '待定',
            trip_type: row['出车方式'] || '公务用车',
            mileage: row['公里数'] || 0,
            amount: row['金额'] || 0,
            remark: row['备注'] || ''
          }]);
        
        if (error) throw error;
        results.success++;
      } catch (error) {
        results.failed++;
        results.errors.push(`第${rowNum}行：${error.message}`);
      }
    }
    
    res.json({ success: true, data: results });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

function formatDateTime(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return dateStr;
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
