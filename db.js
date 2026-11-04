// ====== sql.js 包装层：提供 better-sqlite3 风格的同步 API ======
// 这样 server.js 里的 db.prepare().run/.get/.all 代码几乎不用改
const initSqlJs = require('sql.js');
const fs = require('fs');
const path = require('path');

let _db = null;
let _SQL = null;
const _DB_PATH = process.env.DB_PATH || path.join(__dirname, 'data', 'arc-furnace.db');

function _toObj(row, columns) {
  const obj = {};
  columns.forEach((col, i) => obj[col] = row[i]);
  return obj;
}

function prepare(sql) {
  // 预处理参数占位符：sql.js 也用 ?，和 better-sqlite3 一样
  return {
    run: function(...params) {
      try {
        _db.run(sql, params);
        return { lastInsertRowid: _db.exec('SELECT last_insert_rowid()')[0]?.values?.[0]?.[0] || null, changes: null };
      } catch(e) { throw new Error(e.message); }
    },
    get: function(...params) {
      try {
        const stmt = _db.prepare(sql);
        if (params.length > 0) stmt.bind(params);
        if (stmt.step()) {
          const obj = _toObj(stmt.getAsObject(), stmt.getColumnNames());
          stmt.free();
          return obj;
        }
        stmt.free();
        return undefined;
      } catch(e) { throw new Error(e.message); }
    },
    all: function(...params) {
      try {
        const stmt = _db.prepare(sql);
        if (params.length > 0) stmt.bind(params);
        const results = [];
        while (stmt.step()) {
          results.push(_toObj(stmt.getAsObject(), stmt.getColumnNames()));
        }
        stmt.free();
        return results;
      } catch(e) { throw new Error(e.message); }
    }
  };
}

function exec(sql) {
  return _db.run(sql);
}

function pragma() { /* sql.js 不支持 pragma，忽略 */ }

function save() {
  try {
    fs.mkdirSync(path.dirname(_DB_PATH), { recursive: true });
    const data = _db.export();
    fs.writeFileSync(_DB_PATH, Buffer.from(data));
  } catch(e) { console.error('保存数据库失败:', e.message); }
}

async function init() {
  _SQL = await initSqlJs();
  
  if (fs.existsSync(_DB_PATH)) {
    const data = fs.readFileSync(_DB_PATH);
    _db = new _SQL.Database(data);
    console.log('[DB] 加载已有数据库:', _DB_PATH);
  } else {
    _db = new _SQL.Database();
    console.log('[DB] 创建新数据库');
  }
  
  // 持久化：每 30 秒自动保存一次
  setInterval(save, 30000);
  
  // 进程退出时保存
  process.on('exit', save);
  process.on('SIGINT', () => { save(); process.exit(0); });
  process.on('SIGTERM', () => { save(); process.exit(0); });
  process.on('uncaughtException', (err) => { save(); console.error(err); process.exit(1); });
  
  return { prepare, exec, pragma, save, _db };
}

module.exports = { init };
