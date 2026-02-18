        function showProgress(title, subtitle, steps) {
            document.getElementById('progressTitle').textContent = title;
            document.getElementById('progressSubtitle').textContent = subtitle;

            const stepsContainer = document.getElementById('progressSteps');
            stepsContainer.innerHTML = steps.map((step, index) => `
                <div class="progress-step" id="progressStep${index + 1}">
                    <div class="progress-step-number">${index + 1}</div>
                    <div class="progress-step-text">
                        <div class="progress-step-title">${step.title}</div>
                        <div class="progress-step-desc">${step.desc}</div>
                    </div>
                </div>
            `).join('');

            document.getElementById('progressOverlay').classList.add('show');
            updateProgress(0, 1);
        }

        function hideProgress() {
            document.getElementById('progressOverlay').classList.remove('show');
            document.getElementById('progressPercentage').textContent = '0%';
            document.getElementById('progressBarFill').style.width = '0%';
        }

        function updateProgress(percentage, currentStep) {
            document.getElementById('progressPercentage').textContent = Math.round(percentage) + '%';
            document.getElementById('progressBarFill').style.width = percentage + '%';

            const steps = document.querySelectorAll('.progress-step');
            steps.forEach((step, index) => {
                const stepNum = index + 1;
                if (stepNum < currentStep) {
                    step.classList.remove('active');
                    step.classList.add('completed');
                } else if (stepNum === currentStep) {
                    step.classList.remove('completed');
                    step.classList.add('active');
                } else {
                    step.classList.remove('active', 'completed');
                }
            });
        }

        function delay(ms) {
            return new Promise(resolve => setTimeout(resolve, ms));
        }
        function toggleColumnMapping(show) {
            const div = document.getElementById('columnMappingSection');
            if (show) div.classList.remove('hidden');
            else div.classList.add('hidden');
        }
        function showValidationErrorsPayments(errors) {
            const errorDiv = document.getElementById('validationErrorsPayments');
            const errorList = document.getElementById('errorListPayments');

            errorList.innerHTML = '';
            errors.forEach(error => {
                const li = document.createElement('li');
                li.textContent = error;
                errorList.appendChild(li);
            });

            errorDiv.classList.add('show');
        }

        function hideValidationErrorsPayments() {
            document.getElementById('validationErrorsPayments').classList.remove('show');
        }
        function formatNumber(num) {
            if (num === null || num === undefined || num === '') return '';
            const number = parseFloat(num);
            if (isNaN(number)) return '';
            return number.toLocaleString('uk-UA', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
        }
