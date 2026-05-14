const { app, BrowserWindow, Menu } = require('electron');
const path = require('path');

let mainWindow;

function createWindow() {
    // Создаём окно без лишних элементов
    mainWindow = new BrowserWindow({
        width: 1300,
        height: 900,
        minWidth: 800,
        minHeight: 600,
        icon: path.join(__dirname, 'icon.ico'),
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        },
        title: 'KanjiDB - Японские иероглифы',
        backgroundColor: '#D8CEA9',
        show: false
    });

    // Загружаем HTML
    mainWindow.loadFile('index.html');

    // Убираем стандартное меню
    Menu.setApplicationMenu(null);

    // Показываем окно после загрузки
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    // Закрытие приложения
    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

// Запуск приложения
app.whenReady().then(() => {
    createWindow();
});

// Закрытие на macOS
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});