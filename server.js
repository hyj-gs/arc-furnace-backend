// ====== 电弧炉炼钢过程记录 - 多人协作后端 ======
const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { init: initDb } = require('./db');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '1mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ====== 数据库初始化（async，sql.js 是 WASM 需要异步加载） ======
let db;

async function startup() {
  db = await initDb();
  
  // 确保 data 目录存在
  fs.mkdirSync(path.join(__dirname, 'data'), { recursive: true });
  
  db.exec(`
CREATE TABLE IF NOT EXISTS heats (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  heat_no TEXT NOT NULL UNIQUE,
  steel_type TEXT DEFAULT '30CrNi2MoVA',
  temperature REAL,
  humidity REAL,
  pour_date TEXT,
  shift TEXT,
  created_by TEXT,
  created_at TEXT DEFAULT (datetime('now','localtime')),
  updated_at TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS form_data (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  heat_id INTEGER NOT NULL,
  form_no INTEGER NOT NULL,
  data TEXT NOT NULL,
  filled_by TEXT,
  filled_at TEXT DEFAULT (datetime('now','localtime')),
  UNIQUE(heat_id, form_no)
);

CREATE INDEX IF NOT EXISTS idx_form_heat ON form_data(heat_id);
CREATE INDEX IF NOT EXISTS idx_forms_form ON form_data(form_no);
`);

  app.listen(PORT, () => {
    console.log('========================================');
    console.log('  电弧炉炼钢过程记录 - 多人协作服务');
    console.log('========================================');
    console.log(`  本地访问: http://localhost:${PORT}`);
    console.log(`  局域网:   http://0.0.0.0:${PORT}`);
    console.log(`  数据文件: data/arc-furnace.db`);
    console.log(`  Excel导出: /api/export`);
    console.log('========================================');
  });
}

startup().catch(err => {
  console.error('启动失败:', err);
  process.exit(1);
});

// ====== 6 张表字段模板 ======
// field: 字段名 | label: 显示名 | type: input/number/select/checkbox | unit: 单位 | options: 选项
const FORM_TEMPLATES = [
  {
    no: 1,
    title: '表1 配料工序控制记录',
    header: [
      { key: 'heat_no', label: '炉号' },
      { key: 'steel_type', label: '钢种' },
      { key: 'temperature', label: '温度' },
      { key: 'humidity', label: '湿度' },
      { key: 'duty_unit', label: '责任单位', default: '万迎永安' },
      { key: 'duty_person', label: '责任人', default: '杨进' },
      { key: 'charge_start', label: '装料起止时间' },
      { key: 'total_weight', label: '配料总量', unit: 't', type: 'number' },
      { key: 'actual_weight', label: '实际装料量', unit: 't', type: 'number' },
      { key: 'weighed_weight', label: '过磅称重结果', unit: 't', type: 'number' },
      { key: 'tank_marked', label: '料罐标记', type: 'select', options: ['是', '否'] },
      { key: 'inspector', label: '检查人' },
    ],
    items: [
      {
        id: 1, control: '配料执行',
        standard: '按配料计划钢种和投料量进行配料，不得随意变更。',
        fields: [
          { key: 'plan_steel', label: '计划钢种', type: 'select', options: ['30CrNi2MoVA', '其他'] },
          { key: 'plan_weight', label: '投料量', unit: 'Kg', type: 'number' },
        ]
      },
      {
        id: 2, control: '杂物检查',
        standard: '严禁有色金属(铜、锡、铅、锌)、油污、废铁、砖块、泥块、水泥、粘砂、炉渣等杂物及有害物、爆炸物、封闭容器进入料罐。',
        fields: [
          { key: 'debris_check', label: '检查结果', type: 'select', options: ['无', '有'] },
        ]
      },
      {
        id: 3, control: '过磅称重',
        standard: '严格过磅制，做到重量准确，并将称重结果标记在料罐上。',
        fields: [
          { key: 'weighed_wt', label: '重量', unit: 'Kg', type: 'number' },
          { key: 'tank_mark', label: '标记料罐', type: 'select', options: ['是', '否'] },
        ]
      }
    ]
  },
  {
    no: 2,
    title: '表2 装料工序控制记录',
    header: [
      { key: 'heat_no', label: '炉号' },
      { key: 'steel_type', label: '钢种' },
      { key: 'temperature', label: '温度' },
      { key: 'humidity', label: '湿度' },
      { key: 'duty_team', label: '责任班组', default: '高亮 刁翔' },
      { key: 'duty_person', label: '责任人', default: '高亮 刁翔' },
      { key: 'inspector', label: '检查人' },
    ],
    items: [
      {
        id: 1, control: '装料方式',
        standard: '装料应紧密；轻薄料装罐底，大料装中部，切头小冒口等填空档，小料、中料装上部；严禁渣钢、大料、粘砂铸件冒口等装在上部。',
        fields: [
          { key: 'charge_method', label: '装料方式', type: 'select', options: ['符合要求', '不符合要求'] },
          { key: 'charge_method_note', label: '不符合描述' },
        ]
      },
      {
        id: 2, control: '炉底加石灰',
        standard: '装料前根据炉料和钢种磷含量要求，在过程中加入炉料重量1.5%～4%的石灰和氧化铁皮。',
        fields: [
          { key: 'p_requirement', label: 'P含量要求', type: 'select', options: ['≤0.025%', '其他'] },
          { key: 'lime_weight', label: '石灰重量', unit: 'Kg', type: 'number' },
          { key: 'scale_weight', label: '氧化铁皮重量', unit: 'Kg', type: 'number' },
        ]
      },
      {
        id: 3, control: '加余料时机',
        standard: '加余料时，须在入炉料熔化(70~80)%时装入，装料速度尽可能快。',
        fields: [
          { key: 'remelting_pct', label: '炉料熔化约', unit: '%', type: 'number' },
        ]
      }
    ]
  },
  {
    no: 3,
    title: '表3 熔化与氧化工序控制记录',
    header: [
      { key: 'heat_no', label: '炉号' },
      { key: 'steel_type', label: '钢种' },
      { key: 'temperature', label: '温度' },
      { key: 'humidity', label: '湿度' },
      { key: 'pour_date', label: '浇注日期' },
      { key: 'shift', label: '班次' },
      { key: 'duty_person', label: '责任人', default: '刘瑞 尚伟' },
      { key: 'inspector', label: '检查人' },
    ],
    items: [
      {
        id: 1, control: '熔清样分析',
        standard: '炉料熔化至熔池面上固体炉料低于20%，搅拌确认无大块固体炉料后，取熔清样分析钢水碳、磷等元素。',
        fields: [
          { key: 'melting_C', label: 'C', unit: '%', type: 'number', step: '0.001' },
          { key: 'melting_P', label: 'P', unit: '%', type: 'number', step: '0.001' },
        ]
      },
      {
        id: 2, control: '熔清碳含量',
        standard: '熔清碳含量低于0.50%时，应插电极增碳。',
        fields: [
          { key: 'melting_C_low', label: '熔清碳含量', unit: '%', type: 'number', step: '0.001' },
        ]
      },
      {
        id: 3, control: '扒渣换渣',
        standard: '分析结果磷高于该钢种要求或不能自动流渣时，应及时采取扒渣换渣操作。并加入石灰、萤石、铁矿石或氧化铁皮脱磷造渣。',
        fields: [
          { key: 'slag_P', label: 'P含量', unit: '%', type: 'number', step: '0.001' },
          { key: 'slag_lime', label: '石灰重量', unit: 'Kg', type: 'number' },
          { key: 'slag_fluorspar', label: '萤石重量', unit: 'Kg', type: 'number' },
          { key: 'slag_iron_ore', label: '铁矿石重量', unit: 'Kg', type: 'number' },
        ]
      },
      {
        id: 4, control: '氧化前检测',
        standard: '氧化操作开始前，应检测炉内钢水温度、碳含量、磷含量，满足要求后方可进行氧化操作。',
        fields: [
          { key: 'preox_temp', label: '钢水温度', unit: '℃', type: 'number' },
          { key: 'preox_P', label: 'P含量', unit: '%', type: 'number', step: '0.001' },
          { key: 'preox_C', label: 'C含量', unit: '%', type: 'number', step: '0.001' },
        ]
      },
      {
        id: 5, control: '氧化完毕加锰铁',
        standard: '氧化完毕时，当碳含量≤0.20%时，按要求加入锰铁。',
        fields: [
          { key: 'postox_C', label: 'C含量', unit: '%', type: 'number', step: '0.001' },
          { key: 'postox_FeMn', label: '锰铁重量', unit: 'Kg', type: 'number' },
        ]
      },
      {
        id: 6, control: '合金矿石烘烤',
        standard: '加入炉内的合金、矿石等应经充分烘烤，根据需要准确称量并做好记录。',
        fields: [
          { key: 'bake_temp', label: '烘烤温度', unit: '℃', type: 'number' },
          { key: 'bake_time', label: '烘烤时长', unit: 'min', type: 'number' },
          { key: 'alloys_FeMn_high', label: '高锰', unit: 'Kg', type: 'number' },
          { key: 'alloys_FeCr', label: '高铬', unit: 'Kg', type: 'number' },
          { key: 'alloys_FeMo', label: '钼铁', unit: 'Kg', type: 'number' },
          { key: 'alloys_Ni', label: '镍板', unit: 'Kg', type: 'number' },
          { key: 'alloys_SiAlBa', label: '硅铝钡', unit: 'Kg', type: 'number' },
          { key: 'alloys_RE', label: '稀土', unit: 'Kg', type: 'number' },
        ]
      },
      {
        id: 7, control: '出钢前检测',
        standard: '出钢前应测量碳、磷、硫成分、出钢温度，确定满足要求后方可进行出钢操作。',
        fields: [
          { key: 'tap_C', label: 'C含量', unit: '%', type: 'number', step: '0.001' },
          { key: 'tap_P', label: 'P含量', unit: '%', type: 'number', step: '0.001' },
          { key: 'tap_S', label: 'S含量', unit: '%', type: 'number', step: '0.001' },
          { key: 'tap_temp', label: '钢水温度', unit: '℃', type: 'number' },
        ]
      }
    ]
  },
  {
    no: 4,
    title: '表4 LF精炼工序控制记录',
    header: [
      { key: 'heat_no', label: '炉号' },
      { key: 'steel_type', label: '钢种' },
      { key: 'temperature', label: '温度' },
      { key: 'humidity', label: '湿度' },
      { key: 'pour_date', label: '浇注日期' },
      { key: 'shift', label: '班次' },
      { key: 'duty_person', label: '责任人', default: '刘瑞 尚伟' },
      { key: 'inspector', label: '检查人' },
    ],
    items: [
      {
        id: 1, control: '白渣造渣制度',
        standard: '优化白渣造渣制度，还原期白渣保持时间≥20分钟。',
        fields: [
          { key: 'whiteslag_start', label: '白渣开始时间', unit: '时:分' },
          { key: 'whiteslag_end', label: '白渣结束时间', unit: '时:分' },
        ]
      },
      {
        id: 2, control: '脱氧剂加入',
        standard: '脱氧剂按照吨钢1.5kg小批量多批次加入，均匀撒布在渣面上。',
        fields: [
          { key: 'deoxidizer_1', label: '一次加入重量', unit: 'Kg', type: 'number' },
          { key: 'deoxidizer_2', label: '二次加入重量', unit: 'Kg', type: 'number' },
          { key: 'deoxidizer_3', label: '三次加入重量', unit: 'Kg', type: 'number' },
          { key: 'deoxidizer_4', label: '四次加入重量', unit: 'Kg', type: 'number' },
        ]
      },
      {
        id: 3, control: '喂丝与搅拌',
        standard: '严格控制喂丝速度与吹氩搅拌强度，消除成分偏析。',
        fields: [
          { key: 'argon_flow', label: '吹氩搅拌流量' },
        ]
      },
      {
        id: 4, control: '终脱氧喂丝',
        standard: '终脱氧时根据钢中铝含量进行喂丝。',
        fields: [
          { key: 'Al_content', label: '铝含量', unit: '%', type: 'number', step: '0.001' },
        ]
      },
      {
        id: 5, control: '定氧仪检测氧含量',
        standard: '利用定氧仪检测终脱氧后氧含量。',
        fields: [
          { key: 'O_content', label: '终脱氧后氧含量', unit: 'PPm', type: 'number' },
        ]
      },
      {
        id: 6, control: '吊包出钢',
        standard: '喂丝完成后吊包出钢。',
        fields: [
          { key: 'wire_meter', label: '喂丝量', unit: 'm', type: 'number' },
        ]
      }
    ]
  },
  {
    no: 5,
    title: '表5 VD真空工序控制记录',
    header: [
      { key: 'heat_no', label: '炉号' },
      { key: 'steel_type', label: '钢种' },
      { key: 'temperature', label: '温度' },
      { key: 'humidity', label: '湿度' },
      { key: 'pour_date', label: '浇注日期' },
      { key: 'shift', label: '班次' },
      { key: 'duty_person', label: '责任人', default: '王鹏 赵海' },
      { key: 'inspector', label: '检查人' },
    ],
    items: [
      {
        id: 1, control: '氩气流量控制',
        standard: 'VD过程中应根据不同阶段调整不同的氩气流量，根据透气砖情况灵活调整氩气流量大小。',
        fields: [
          { key: 'argon_evac', label: '抽真空氩气流量' },
          { key: 'argon_hold', label: '保真空氩气流量' },
          { key: 'argon_soft', label: '软吹氩气流量' },
        ]
      },
      {
        id: 2, control: '破真空测温取样',
        standard: '破真空后根据要求进行测温、取样操作。',
        fields: [
          { key: 'break_temp', label: '温度', unit: '℃', type: 'number' },
        ]
      },
      {
        id: 3, control: '温度成分确认',
        standard: '温度、成分合格后方可出钢。',
        fields: [
          { key: 'confirm_temp', label: '温度', unit: '℃', type: 'number' },
          { key: 'comp_C', label: 'C', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_Si', label: 'Si', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_Mn', label: 'Mn', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_P', label: 'P', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_S', label: 'S', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_Cr', label: 'Cr', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_Ni', label: 'Ni', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_Mo', label: 'Mo', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_Cu', label: 'Cu', unit: '%', type: 'number', step: '0.001' },
          { key: 'comp_V', label: 'V', unit: '%', type: 'number', step: '0.001' },
        ]
      },
      {
        id: 4, control: '软吹',
        standard: '记录软吹时长',
        fields: [
          { key: 'soft_start', label: '软吹开始时间', unit: '时:分' },
          { key: 'soft_end', label: '软吹结束时间', unit: '时:分' },
        ]
      }
    ]
  },
  {
    no: 6,
    title: '表6 浇注与脱模工序控制记录',
    header: [
      { key: 'heat_no', label: '炉号' },
      { key: 'steel_type', label: '钢种' },
      { key: 'temperature', label: '温度' },
      { key: 'humidity', label: '湿度' },
      { key: 'pour_date', label: '浇注日期' },
      { key: 'shift', label: '班次' },
      { key: 'duty_person', label: '责任人', default: '肖洪斌' },
      { key: 'inspector', label: '检查人' },
    ],
    items: [
      {
        id: 1, control: '浇注前检查',
        standard: '浇注前应对底盘、中注管、钢锭模、冒口进行检查，合格后方可开浇。',
        fields: [
          { key: 'chk_baseplate', label: '底盘', type: 'select', options: ['合格', '不合格'] },
          { key: 'chk_tundish', label: '中注管', type: 'select', options: ['合格', '不合格'] },
          { key: 'chk_mould', label: '钢锭模', type: 'select', options: ['合格', '不合格'] },
          { key: 'chk_riser', label: '冒口', type: 'select', options: ['合格', '不合格'] },
        ]
      },
      {
        id: 2, control: '保护浇注',
        standard: '浇注中采用蜜蜡圈做好钢水保护浇注，减少钢液在浇注过程裸露吸气。',
        fields: [
          { key: 'protect_pour', label: '蜜蜡圈使用', type: 'select', options: ['使用', '未使用'] },
        ]
      },
      {
        id: 3, control: '脱模后缓冷',
        standard: '规范脱模后缓冷，高碳钢及合金钢锭进坑缓冷。',
        fields: [
          { key: 'slow_cool', label: '是否进坑', type: 'select', options: ['是', '否'] },
        ]
      },
      {
        id: 4, control: '退火处理',
        standard: '高合金钢锭热脱后进行退火。',
        fields: [
          { key: 'anneal', label: '是否退火', type: 'select', options: ['是', '否'] },
        ]
      }
    ]
  }
];

// ====== API 路由 ======

// 模板接口
app.get('/api/templates', (req, res) => {
  res.json(FORM_TEMPLATES.map(t => ({
    no: t.no,
    title: t.title,
    header: t.header,
    items: t.items.map(i => ({ id: i.id, control: i.control, standard: i.standard, fields: i.fields }))
  })));
});

// 炉次列表（含填写进度）
app.get('/api/heats', (req, res) => {
  const rows = db.prepare(`
    SELECT h.*, 
      (SELECT GROUP_CONCAT(form_no) FROM form_data WHERE heat_id = h.id) as filled_forms
    FROM heats h ORDER BY h.created_at DESC
  `).all();
  
  res.json(rows.map(r => ({
    id: r.id,
    heat_no: r.heat_no,
    steel_type: r.steel_type,
    temperature: r.temperature,
    humidity: r.humidity,
    pour_date: r.pour_date,
    shift: r.shift,
    created_by: r.created_by,
    created_at: r.created_at,
    progress: r.filled_forms ? r.filled_forms.split(',').map(Number) : []
  })));
});

// 创建炉次
app.post('/api/heats', (req, res) => {
  const { heat_no, steel_type, temperature, humidity, pour_date, shift, created_by } = req.body;
  
  // 自动生成炉号
  let finalHeatNo = heat_no;
  if (!finalHeatNo) {
    const now = new Date();
    const seq = db.prepare(`SELECT COUNT(*) + 1 as seq FROM heats WHERE steel_type = ?`).get(steel_type || '30CrNi2MoVA').seq;
    finalHeatNo = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}-${String(seq).padStart(3,'0')}`;
  }
  
  try {
    const info = db.prepare(`INSERT INTO heats (heat_no, steel_type, temperature, humidity, pour_date, shift, created_by)
      VALUES (?, ?, ?, ?, ?, ?, ?)`).run(finalHeatNo, steel_type || '30CrNi2MoVA', 
      temperature || null, humidity || null, pour_date || null, shift || null, created_by || '匿名');
    res.json({ id: info.lastInsertRowid, heat_no: finalHeatNo });
  } catch(e) {
    if (e.message.includes('UNIQUE')) {
      res.status(400).json({ error: '炉号已存在' });
    } else {
      res.status(500).json({ error: e.message });
    }
  }
});

// 获取单个炉次
app.get('/api/heats/:id', (req, res) => {
  const heat = db.prepare(`SELECT * FROM heats WHERE id = ?`).get(req.params.id);
  if (!heat) return res.status(404).json({ error: '炉次不存在' });
  
  const forms = db.prepare(`SELECT form_no, data, filled_by, filled_at FROM form_data WHERE heat_id = ?`).all(req.params.id);
  const formMap = {};
  forms.forEach(f => { formMap[f.form_no] = { data: JSON.parse(f.data), filled_by: f.filled_by, filled_at: f.filled_at }; });
  
  res.json({ ...heat, forms: formMap });
});

// 更新炉次基础信息
app.put('/api/heats/:id', (req, res) => {
  const { heat_no, steel_type, temperature, humidity, pour_date, shift } = req.body;
  db.prepare(`UPDATE heats SET heat_no=?, steel_type=?, temperature=?, humidity=?, pour_date=?, shift=?, updated_at=datetime('now','localtime') WHERE id=?`)
    .run(heat_no, steel_type, temperature, humidity, pour_date, shift, req.params.id);
  res.json({ ok: true });
});

// 保存/更新某张表的填写数据
app.post('/api/heats/:id/forms/:formNo', (req, res) => {
  const heatId = req.params.id;
  const formNo = parseInt(req.params.formNo);
  const { data, filled_by } = req.body;
  
  // 先确保炉次存在
  const heat = db.prepare('SELECT * FROM heats WHERE id = ?').get(heatId);
  if (!heat) return res.status(404).json({ error: '炉次不存在' });
  
  const existing = db.prepare('SELECT id FROM form_data WHERE heat_id = ? AND form_no = ?').get(heatId, formNo);
  
  if (existing) {
    db.prepare(`UPDATE form_data SET data = ?, filled_by = ?, filled_at = datetime('now','localtime') 
      WHERE id = ?`).run(JSON.stringify(data), filled_by || null, existing.id);
  } else {
    db.prepare(`INSERT INTO form_data (heat_id, form_no, data, filled_by) VALUES (?, ?, ?, ?)`)
      .run(heatId, formNo, JSON.stringify(data), filled_by || null);
  }
  
  db.prepare(`UPDATE heats SET updated_at = datetime('now','localtime') WHERE id = ?`).run(heatId);
  res.json({ ok: true });
});

// 删除炉次
app.delete('/api/heats/:id', (req, res) => {
  db.prepare('DELETE FROM form_data WHERE heat_id = ?').run(req.params.id);
  db.prepare('DELETE FROM heats WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

// 统计
app.get('/api/stats', (req, res) => {
  const heatCount = db.prepare('SELECT COUNT(*) as c FROM heats').get().c;
  const filledCount = db.prepare('SELECT COUNT(*) as c FROM form_data').get().c;
  const uniqueFillers = db.prepare("SELECT COUNT(DISTINCT filled_by) as c FROM form_data WHERE filled_by IS NOT NULL AND filled_by != ''").get().c;
  
  // 每张表的填写数量
  const formStats = db.prepare(`SELECT form_no, COUNT(*) as c FROM form_data GROUP BY form_no ORDER BY form_no`).all();
  
  res.json({ heatCount, filledCount, uniqueFillers, formStats });
});

// 导出 Excel（汇总所有炉次所有表）
app.get('/api/export', (req, res) => {
  const heats = db.prepare('SELECT * FROM heats ORDER BY created_at DESC').all();
  if (heats.length === 0) return res.status(404).json({ error: '没有数据' });
  
  const wb = XLSX.utils.book_new();
  
  // Sheet 1: 炉次总览
  const overviewData = [['炉号', '钢种', '温度', '湿度', '浇注日期', '班次', '创建人', '创建时间', '表1', '表2', '表3', '表4', '表5', '表6', '是否齐全']];
  heats.forEach(h => {
    const progress = db.prepare('SELECT form_no FROM form_data WHERE heat_id = ?').all(h.id).map(r => r.form_no);
    const row = [h.heat_no, h.steel_type, h.temperature || '', h.humidity || '', h.pour_date || '', h.shift || '', h.created_by || '', h.created_at];
    for (let i = 1; i <= 6; i++) row.push(progress.includes(i) ? '✅' : '❌');
    row.push(progress.length === 6 ? '✅ 齐全' : `${progress.length}/6`);
    overviewData.push(row);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(overviewData), '炉次总览');
  
  // Sheet 2-7: 每张表的所有数据
  FORM_TEMPLATES.forEach(tpl => {
    const rows = [];
    // 构建表头
    const header = ['炉号', '钢种', '钢水温度', '湿度', '浇注日期', '班次', '填写人', '填写时间'];
    tpl.header.forEach(h => {
      if (['heat_no', 'steel_type', 'temperature', 'humidity', 'pour_date', 'shift'].includes(h.key)) return;
      header.push(h.label);
    });
    tpl.items.forEach(item => {
      header.push(`【${item.id}】${item.control}`);
      item.fields.forEach(f => {
        header.push(`  ${f.label}${f.unit ? '(' + f.unit + ')' : ''}`);
      });
    });
    rows.push(header);
    
    heats.forEach(h => {
      const fd = db.prepare('SELECT * FROM form_data WHERE heat_id = ? AND form_no = ?').get(h.id, tpl.no);
      const data = fd ? JSON.parse(fd.data) : {};
      const row = [h.heat_no, h.steel_type, data.temperature || h.temperature || '', data.humidity || h.humidity || '', 
                   data.pour_date || h.pour_date || '', data.shift || h.shift || '', fd?.filled_by || '', fd?.filled_at || ''];
      tpl.header.forEach(hdr => {
        if (['heat_no', 'steel_type', 'temperature', 'humidity', 'pour_date', 'shift'].includes(hdr.key)) return;
        row.push(data[hdr.key] ?? '');
      });
      tpl.items.forEach(item => {
        row.push(''); // 控制项目名
        item.fields.forEach(f => row.push(data[f.key] ?? ''));
      });
      rows.push(row);
    });
    
    const sheetName = tpl.title.replace(/^表\d+\s*/, '').slice(0, 28);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
  });
  
  // 文件名
  const now = new Date().toISOString().slice(0, 10);
  const filename = `电弧炉炼钢记录_${now}.xlsx`;
  
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
  res.send(buffer);
});

// 健康检查
app.get('/api/health', (req, res) => res.json({ status: 'ok', time: new Date().toISOString() }));
