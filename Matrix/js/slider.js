// ========== ВСЕ СКРИПТЫ ==========
        // Получаем элементы
        const heightSlider = document.getElementById('heightSlider');
        const heightInput = document.getElementById('heightInput');
        const lengthSlider = document.getElementById('lengthSlider');
        const lengthInput = document.getElementById('lengthInput');

        // Функция обновления высоты
        function updateHeight(value) {
            const num = Number(value);
            if (isNaN(num)) return;
            const clamped = Math.min(100, Math.max(1, num));
            
            heightSlider.value = clamped;
            heightInput.value = clamped;
        }

        // Функция обновления длины
        function updateLength(value) {
            const num = Number(value);
            if (isNaN(num)) return;
            const clamped = Math.min(100, Math.max(1, num));
            
            lengthSlider.value = clamped;
            lengthInput.value = clamped;
        }

        // События для высоты
        heightSlider.addEventListener('input', function() {
            updateHeight(this.value);
        });

        heightInput.addEventListener('input', function() {
            if (this.value === '' || isNaN(Number(this.value))) return;
            updateHeight(this.value);
        });

        heightInput.addEventListener('blur', function() {
            const val = Number(this.value);
            if (isNaN(val)) {
                updateHeight(heightSlider.value);
            } else {
                updateHeight(Math.round(val));
            }
        });

        // События для длины
        lengthSlider.addEventListener('input', function() {
            updateLength(this.value);
        });

        lengthInput.addEventListener('input', function() {
            if (this.value === '' || isNaN(Number(this.value))) return;
            updateLength(this.value);
        });

        lengthInput.addEventListener('blur', function() {
            const val = Number(this.value);
            if (isNaN(val)) {
                updateLength(lengthSlider.value);
            } else {
                updateLength(Math.round(val));
            }
        });