const express = require('express');
const cors = require('cors');
const sqlite3 = require('sqlite3').verbose();
const ExcelJS = require('exceljs');
const fs = require('node:fs');
const path = require('node:path');

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const databasePath = path.join(__dirname, 'database', 'supervision.db');
const plantillaPath = path.join(__dirname, 'plantilla', 'plantilla_supervision.xlsm');
const fields = [
  'oficina_regional',
  'mes_presenta',
  'accion',
  'accion_especifica',
  'sub_accion',
  'tipo_agente',
  'actividad',
  'provincia',
  'distrito',
  'tipo_transporte',
  'num_supervisores',
  'empresa_supervisora',
  'calidad_entregable',
  'expediente_siged',
  'carta_linea',
  'observaciones'
];

fs.mkdirSync(path.dirname(databasePath), { recursive: true });

app.use(cors());
app.use(express.json({ limit: '100kb' }));
app.use(express.static(path.join(__dirname, '..', 'frontend')));

const db = new sqlite3.Database(databasePath, (error) => {
  if (error) {
    console.error('Error al conectar a la BD:', error.message);
    return;
  }

  console.log('Conectado a la base de datos SQLite.');
});

db.exec(fs.readFileSync(path.join(__dirname, 'database', 'schema.sql'), 'utf8'), (error) => {
  if (error) {
    console.error('Error al crear las tablas de la base de datos:', error.message);
    process.exitCode = 1;
    return;
  }

  app.listen(PORT, () => {
    console.log(`Servidor ejecutándose en http://localhost:${PORT}`);
  });
});

app.get('/api/health', (_request, response) => {
  response.json({ status: 'ok' });
});

app.get('/api/registros', (_request, response) => {
  db.all('SELECT * FROM registros ORDER BY id DESC', [], (error, rows) => {
    if (error) {
      return response.status(500).json({ error: error.message });
    }

    return response.json(rows);
  });
});

app.post('/api/registros', (request, response) => {
  const data = request.body;

  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    return response.status(400).json({ error: 'El cuerpo de la solicitud debe ser un objeto JSON.' });
  }

  const invalidField = fields.find((field) => {
    if (field === 'num_supervisores') {
      return data[field] !== undefined &&
        data[field] !== null &&
        data[field] !== '' &&
        (!Number.isInteger(Number(data[field])) || Number(data[field]) < 0);
    }

    return data[field] !== undefined &&
      data[field] !== null &&
      typeof data[field] !== 'string';
  });

  if (invalidField) {
    return response.status(400).json({ error: `El campo ${invalidField} no es válido.` });
  }

  const sql = `INSERT INTO registros (${fields.join(', ')})
    VALUES (${fields.map(() => '?').join(', ')})`;
  const params = fields.map((field) => {
    if (data[field] === undefined || data[field] === '') {
      return null;
    }

    return field === 'num_supervisores' ? Number(data[field]) : data[field];
  });

  db.run(sql, params, function (error) {
    if (error) {
      return response.status(400).json({ error: error.message });
    }

    return response.status(201).json({
      id: this.lastID,
      message: 'Registro guardado exitosamente.'
    });
  });
});

app.delete('/api/registros/:id', (request, response) => {
  const id = Number(request.params.id);
  if (!Number.isSafeInteger(id) || id < 1) {
    return response.status(400).json({ error: 'El identificador del registro no es válido.' });
  }

  db.run('DELETE FROM registros WHERE id = ?', id, function (error) {
    if (error) {
      return response.status(400).json({ error: error.message });
    }
    if (this.changes === 0) {
      return response.status(404).json({ error: 'No se encontró el registro.' });
    }

    return response.json({ message: 'Registro eliminado.', changes: this.changes });
  });
});

app.get('/api/exportar-excel', async (_request, response, next) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(plantillaPath);
    const worksheet = workbook.getWorksheet('Registro');

    if (!worksheet) {
      throw new Error('La plantilla no contiene la hoja "Registro".');
    }

    db.all('SELECT * FROM registros ORDER BY id', [], async (error, rows) => {
      if (error) {
        return response.status(500).json({ error: error.message });
      }

      try {
        const headers = [
          'Oficina regional',
          'Mes de presentación',
          'Acción',
          'Acción específica',
          'Subacción específica',
          'Tipo de agente',
          'Actividad',
          'Provincia',
          'Distrito',
          'Tipo de transporte',
          'N.º supervisores',
          'Empresa supervisora',
          'Calidad del entregable',
          'Expediente SIGED',
          'Carta / línea',
          'Observaciones'
        ];
        const columnWidths = [19, 20, 24, 28, 25, 20, 25, 19, 22, 21, 16, 28, 23, 23, 19, 40];

        worksheet.spliceRows(1, worksheet.rowCount);
        worksheet.columns = columnWidths.map((width) => ({ width }));
        worksheet.mergeCells('A1:P1');
        worksheet.getCell('A1').value = 'REPORTE DE SUPERVISIÓN DE HIDROCARBUROS';
        worksheet.mergeCells('A2:P2');
        worksheet.getCell('A2').value = 'Oficina Regional de Huánuco · Registro de acciones de supervisión';
        worksheet.getCell('A4').value = 'Total de registros';
        worksheet.getCell('B4').value = rows.length;
        worksheet.getCell('D4').value = 'Fecha de exportación';
        worksheet.getCell('E4').value = new Date();
        worksheet.getCell('E4').numFmt = 'dd/mm/yyyy hh:mm';
        worksheet.getRow(6).values = headers;

        rows.forEach((record) => {
          worksheet.addRow(fields.map((field) => record[field] ?? ''));
        });

        const lastRow = worksheet.rowCount;
        const title = worksheet.getRow(1);
        title.height = 34;
        title.getCell(1).font = { name: 'Aptos Display', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
        title.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
        title.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF143747' } };

        const subtitle = worksheet.getRow(2);
        subtitle.height = 24;
        subtitle.getCell(1).font = { name: 'Aptos', size: 11, color: { argb: 'FFFFFFFF' } };
        subtitle.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
        subtitle.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A6A64' } };

        for (const address of ['A4', 'D4']) {
          const cell = worksheet.getCell(address);
          cell.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FF31525E' } };
        }
        for (const address of ['B4', 'E4']) {
          const cell = worksheet.getCell(address);
          cell.font = { name: 'Aptos', size: 10, color: { argb: 'FF172B3A' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF3F1' } };
        }

        const header = worksheet.getRow(6);
        header.height = 32;
        header.eachCell((cell) => {
          cell.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF176B61' } };
          cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
          cell.border = {
            bottom: { style: 'medium', color: { argb: 'FF104E48' } }
          };
        });

        for (let rowNumber = 7; rowNumber <= lastRow; rowNumber += 1) {
          const row = worksheet.getRow(rowNumber);
          row.height = 30;
          row.eachCell({ includeEmpty: true }, (cell) => {
            cell.font = { name: 'Aptos', size: 10, color: { argb: 'FF263E49' } };
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: rowNumber % 2 === 0 ? 'FFF0F6F5' : 'FFFFFFFF' }
            };
            cell.alignment = {
              vertical: 'middle',
              horizontal: 'left',
              wrapText: true
            };
            cell.border = {
              bottom: { style: 'hair', color: { argb: 'FFD8E3E5' } }
            };
          });
          row.getCell(11).alignment = { vertical: 'middle', horizontal: 'center' };
        }

        worksheet.autoFilter = `A6:P${Math.max(lastRow, 6)}`;
        worksheet.views = [{ state: 'frozen', ySplit: 6, showGridLines: false }];
        worksheet.pageSetup = {
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          paperSize: 9
        };
        worksheet.pageSetup.printTitlesRow = '1:6';
        worksheet.pageSetup.printArea = `A1:P${Math.max(lastRow, 6)}`;

        const buffer = await workbook.xlsx.writeBuffer();
        response.setHeader(
          'Content-Type',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        response.setHeader(
          'Content-Disposition',
          'attachment; filename=Reporte_Supervision_Hidrocarburos.xlsx'
        );
        return response.send(Buffer.from(buffer));
      } catch (writeError) {
        return next(writeError);
      }
    });
  } catch (error) {
    return next(error);
  }
});

app.use((error, _request, response, _next) => {
  console.error('Error al procesar la solicitud:', error);
  if (response.headersSent) {
    return response.end();
  }

  return response.status(500).json({
    error: `Error al procesar la solicitud: ${error.message}`
  });
});

process.on('SIGINT', () => {
  db.close((error) => {
    if (error) {
      console.error('Error al cerrar la base de datos:', error.message);
      process.exitCode = 1;
    }
    process.exit();
  });
});
