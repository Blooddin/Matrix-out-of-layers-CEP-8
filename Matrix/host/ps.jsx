// ============================================================
// ГЕНЕРАТОР МАТРИЦЫ С СОЗДАНИЕМ НОВОЙ ПАПКИ
// ============================================================

// Главная функция: создает матрицу в новой папке
function generateMatrixInNewFolder(p, m) {
    try {
        // 1. Проверяем, есть ли открытый документ
        if (app.documents.length === 0) {
            return "❌ Нет открытого документа";
        }
        
        var doc = app.activeDocument;
        
        // 2. Ищем папку (группу) с названием "Matrix"
        var matrixFolder = null;
        for (var i = 0; i < doc.layerSets.length; i++) {
            if (doc.layerSets[i].name === "Matrix") {
                matrixFolder = doc.layerSets[i];
                break;
            }
        }
        
        if (!matrixFolder) {
            return "❌ Папка 'Matrix' не найдена!\n" +
                   "Создайте папку и поместите в нее 6 слоев-шаблонов";
        }
        
        // 3. Проверяем, что в папке есть хотя бы 6 слоев
        var sourceLayers = [];
        for (var i = 0; i < matrixFolder.artLayers.length; i++) {
            sourceLayers.push(matrixFolder.artLayers[i]);
        }
        
        if (sourceLayers.length < 6) {
            return "❌ В папке 'Matrix' нужно минимум 6 слоев (сейчас: " + sourceLayers.length + ")";
        }
        
        // 4. Берем только первые 6 слоев
        sourceLayers = sourceLayers.slice(0, 6);
        
        // 5. Определяем размеры одного слоя
        var layerWidth = -1;
        var layerHeight = -1;
        try {
            for(var i = 0; i < sourceLayers.length; i++)
            {
                var bounds = sourceLayers[i].bounds;
                var layerWidthTemp = bounds[2].value - bounds[0].value;
                var layerHeightTemp = bounds[3].value - bounds[1].value;
                
                if(layerWidthTemp > layerWidth){
                    layerWidth = layerWidthTemp;
                }
                if(layerHeightTemp > layerHeight){
                    layerHeight = layerHeightTemp;
                }
            }
            
        } catch (e) {
            // Используем значения по умолчанию
        }
        if(layerWidth == -1 || layerHeight == -1)
        {
            return "❌ Программа не может посчитать границы элемента";
        }
        
        // 6. Создаем новую папку для результата
        var resultFolder = doc.layerSets.add();
        // Генерируем уникальное имя
        var folderName = "MatrixResult_" + getTimestamp();
        resultFolder.name = folderName;
        
        // 7. Создаем матрицу случайных чисел
        var matrix = [];
        for (var i = 0; i < p; i++) {
            var row = [];
            for (var j = 0; j < m; j++) {
                row.push(Math.floor(Math.random() * 6) + 1);
            }
            matrix.push(row);
        }
        
        // 8. Размещаем дубликаты в новой папке
        var spacingW = 0;
        var spacingH = 0;
        var placed = 0;
        var errors = [];
        
        // Сохраняем текущий активный слой
        var originalActiveLayer = doc.activeLayer;
        
        for (var i = 0; i < p; i++) {
            for (var j = 0; j < m; j++) {
                var layerIndex = matrix[i][j] - 1;
                var sourceLayer = sourceLayers[layerIndex];
                
                if (!sourceLayer) {
                    errors.push("Слой #" + (layerIndex + 1) + " не найден");
                    continue;
                }
                
                try {
                    // Дублируем слой
                    var newLayer = sourceLayer.duplicate(resultFolder, ElementPlacement.PLACEATEND);
                    newLayer.name = "Слой_" + (i+1) + "_" + (j+1);

                    
                    // Вычисляем позицию
                    var xPos = i * (layerWidth + spacingW);
                    var yPos = j * (layerHeight + spacingH);
                    
                    
                    // Перемещаем слой
                    var bounds = newLayer.bounds;
                    var dx = xPos - bounds[0].value;
                    var dy = yPos - bounds[1].value;
                    dx-=(bounds[2].value - bounds[0].value - layerWidth)/2;
                    dy-=(bounds[3].value - bounds[1].value - layerHeight)/2;
                    newLayer.translate(dx, dy);
                    
                    placed++;
                } catch (e) {
                    errors.push("Ошибка [" + i + "][" + j + "]: " + e.toString());
                }
            }
        }
        
        // Восстанавливаем активный слой
        try {
            doc.activeLayer = originalActiveLayer;
        } catch (e) {
            // Игнорируем
        }
        
       

        // 9. Формируем результат
        var result = "✅ Матрица " + p + "x" + m + " создана!\n";
        result += "Источник: папка 'Matrix' (" + sourceLayers.length + " слоев)\n";
        result += "Результат: папка '" + folderName + "'\n";
        result += "Размещено слоев: " + placed + " из " + (p * m) + "\n";
        
        if (errors.length > 0) {
            result += "\n⚠️ Ошибки:\n" + errors.join("\n");
        }
        
        result += "\n\n📊 Матрица чисел:\n";
        result += "--------------------------------\n";
        for (var i = 0; i < p; i++) {
            result += "  ";
            for (var j = 0; j < m; j++) {
                result += matrix[i][j] + " ";
            }
            result += "\n";
        }
        result += "--------------------------------";
        
        return result;
        
    } catch (error) {
        return "❌ Ошибка: " + error.toString() + "\n\n" + error.stack;
    }
}

// ============================================================
// ДОПОЛНИТЕЛЬНЫЕ ФУНКЦИИ
// ============================================================

// Получить текущую дату/время для уникального имени
function getTimestamp() {
    var now = new Date();
    var year = now.getFullYear();
    var month = padZero(now.getMonth() + 1);
    var day = padZero(now.getDate());
    var hours = padZero(now.getHours());
    var minutes = padZero(now.getMinutes());
    var seconds = padZero(now.getSeconds());
    return year + month + day + "_" + hours + minutes + seconds;
}

function padZero(num) {
    return num < 10 ? "0" + num : num;
}

// Получить информацию о папке Matrix
function getMatrixFolderInfo() {
    try {
        if (app.documents.length === 0) {
            return "❌ Нет открытого документа";
        }
        
        var doc = app.activeDocument;
        
        // Ищем папку Matrix
        var matrixFolder = null;
        for (var i = 0; i < doc.layerSets.length; i++) {
            if (doc.layerSets[i].name === "Matrix") {
                matrixFolder = doc.layerSets[i];
                break;
            }
        }
        
        if (!matrixFolder) {
            return "❌ Папка 'Matrix' не найдена!\n\n" +
                   "Чтобы создать:\n" +
                   "1. Нажмите 'Создать новую группу' в панели слоев\n" +
                   "2. Назовите ее 'Matrix'\n" +
                   "3. Поместите в нее 6 слоев-шаблонов";
        }
        
        var layers = matrixFolder.artLayers;
        var result = "📁 Папка 'Matrix' найдена!\n";
        result += "----------------------------------------\n";
        result += "Слоев в папке: " + layers.length + "\n";
        result += "Используется: " + Math.min(layers.length, 6) + " из них\n\n";
        
        var count = Math.min(layers.length, 6);
        for (var i = 0; i < count; i++) {
            var layer = layers[i];
            result += (i + 1) + ". " + layer.name;
            result += layer.visible ? " 👁️" : " 🚫";
            
            try {
                var bounds = layer.bounds;
                var w = bounds[2].value - bounds[0].value;
                var h = bounds[3].value - bounds[1].value;
                result += " (" + w.toFixed(0) + "x" + h.toFixed(0) + " px)";
            } catch (e) {
                result += " (размер неизвестен)";
            }
            result += "\n";
        }
        
        if (layers.length > 6) {
            result += "\n⚠️ В папке больше 6 слоев. Используются только первые 6.";
        }
        
        return result;
    } catch (e) {
        return "❌ Ошибка: " + e.toString();
    }
}

// Получить список всех папок MatrixResult
function getResultFolders() {
    try {
        if (app.documents.length === 0) {
            return "❌ Нет открытого документа";
        }
        
        var doc = app.activeDocument;
        var result = "📁 Папки с результатами:\n";
        result += "----------------------------------------\n";
        
        var found = false;
        for (var i = 0; i < doc.layerSets.length; i++) {
            var folder = doc.layerSets[i];
            if (folder.name.indexOf("MatrixResult_") === 0) {
                var layersCount = folder.artLayers.length;
                result += "  📂 " + folder.name + " (" + layersCount + " слоев)\n";
                found = true;
            }
        }
        
        if (!found) {
            result += "  (нет папок с результатами)";
        }
        result += "\n----------------------------------------";
        
        return result;
    } catch (e) {
        return "❌ Ошибка: " + e.toString();
    }
}

// Удалить все папки MatrixResult
function clearResultFolders() {
    try {
        if (app.documents.length === 0) {
            return "❌ Нет открытого документа";
        }
        
        var doc = app.activeDocument;
        var deleted = 0;
        var i = doc.layerSets.length - 1;
        
        while (i >= 0) {
            var folder = doc.layerSets[i];
            if (folder.name.indexOf("MatrixResult_") === 0) {
                folder.remove();
                deleted++;
            }
            i--;
        }
        
        return "✅ Удалено " + deleted + " папок с результатами";
    } catch (e) {
        return "❌ Ошибка: " + e.toString();
    }
}

// Генерация матрицы чисел
function generateNumberMatrix(p, m) {
    try {
        var matrix = [];
        for (var i = 0; i < p; i++) {
            var row = [];
            for (var j = 0; j < m; j++) {
                row.push(Math.floor(Math.random() * 6) + 1);
            }
            matrix.push(row);
        }
        
        var result = "📊 Матрица " + p + "x" + m + ":\n";
        result += "--------------------------------\n";
        for (var i = 0; i < p; i++) {
            result += "  ";
            for (var j = 0; j < m; j++) {
                result += matrix[i][j] + " ";
            }
            result += "\n";
        }
        result += "--------------------------------";
        
        return result;
    } catch (e) {
        return "❌ Ошибка: " + e.toString();
    }
}