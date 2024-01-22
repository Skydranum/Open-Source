const express = require('express');
const mysql = require('mysql')
const cors = require('cors')
const bodyParser = require('body-parser');
const multer = require('multer');

const app = express();
const port = 3001;
app.use(cors());
app.use(bodyParser.json());

// Conexão com Servidor MySQL
const pool = mysql.createConnection({
  connectionLimit: 10, 
  host: 'X',
  user: 'X',
  password: '451jA*FXKApu4-y',
  database: 'X'
});

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, './uploads/');
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ storage: storage });

app.post('/upload', upload.single('pdfFile'), (req, res, next) => {
  console.log("Upload endpoint hit!");
  if (req.file) {
    console.log("File received:", req.file.filename);
  } else {
    console.log("No file received");
  }
  res.json({ filePath: `/uploads/${req.file.filename}` });
});

app.use('/uploads', express.static('uploads'));

app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).send('Something broke!');
});

app.get('/sondagens', (req, res) => {
  pool.query('SELECT id, CoordenadasX, CoordenadasY, Nome, Descricao, Nomeclatura, PDF, Tag1, Tag2, Tag3, Tag4 FROM sondagens', (error, results) => { // Adicione , PDF aqui
    if (error) {
      return res.status(500).json({ error });
    }
    res.json(results);
  });
});

app.post('/sondagens', (req, res) => {
  const { CoordenadasX, CoordenadasY, Nome, Descricao, Nomeclatura, filePath, Tag1, Tag2, Tag3, Tag4 } = req.body;
  const query = 'INSERT INTO sondagens (CoordenadasX, CoordenadasY, Nome, Descricao, Nomeclatura, PDF, Tag1, Tag2, Tag3, Tag4) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)';

  pool.query(query, [CoordenadasX, CoordenadasY, Nome, Descricao, Nomeclatura, filePath, Tag1, Tag2, Tag3, Tag4], (error, results) => {
    if (error) {
      console.error("Error while inserting data:", error);
      return res.status(500).json({ error });
    }
    res.json({ success: true, message: 'Data inserted successfully!' });
  });
});

app.delete('/sondagens/:id', (req, res) => {
  const { id } = req.params;
  pool.query('DELETE FROM sondagens WHERE id = ?', [id], (error, results) => {
    if (error) {
      console.error("Error during delete:", error);
      return res.status(500).json({ error });
    }
    res.json({ success: true, message: 'Data deleted successfully!' });
  });
});

app.put('/sondagens/:id', upload.single('pdfFile'), (req, res) => {
  const { id } = req.params;
  const { CoordenadasX, CoordenadasY, Nome, Descricao, Nomeclatura, Tag1, Tag2, Tag3, Tag4 } = req.body;

  let query = 'UPDATE sondagens SET ';
  let data = [];

  if (CoordenadasX && CoordenadasY) {
    query += 'CoordenadasX = ?, CoordenadasY = ?, ';
    data.push(CoordenadasX, CoordenadasY);
  }

  query += 'Nome = ?, Descricao = ?, Nomeclatura = ?';
  data.push(Nome, Descricao, Nomeclatura);

  // Se um novo arquivo foi enviado, atualize a coluna PDF
  if (req.file) {
    query += ', PDF = ?';
    data.push(`/uploads/${req.file.filename}`);
  }

  if (Tag1) {
    query += ', Tag1 = ?';
    data.push(Tag1);
  }
  if (Tag2) {
    query += ', Tag2 = ?';
    data.push(Tag2);
  }
  if (Tag3) {
    query += ', Tag3 = ?';
    data.push(Tag3);
  }
  if (Tag4) {
    query += ', Tag4 = ?';
    data.push(Tag4);
  }

  query += ' WHERE id = ?';
  data.push(id);

  pool.query(query, data, (error, results) => {
    if (error) {
      console.error("Error during update:", error);
      return res.status(500).json({ error });
    }
    res.json({ success: true, message: 'Data updated successfully!' });
  });
});

app.listen(port, () => {
  console.log(`Server started on http://localhost:${port}`);
});