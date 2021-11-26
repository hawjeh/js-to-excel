import express from 'express';
import path from 'path';
const __dirname = path.resolve();

const app = express();
const port = 8080;
app.use(express.json());
app.use(express.static(path.join(__dirname, './dist')));

app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, './src/test/index.html'));
});

app.listen(port, () => {
    console.log(`Server listening on the port::${port}`);
});
