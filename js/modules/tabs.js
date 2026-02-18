        function switchTab(tabName) {
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.remove('active');
            });

            // Remove active class from all tab buttons
            document.querySelectorAll('.tab').forEach(btn => {
                btn.classList.remove('active');
            });

            // Show selected tab
            if (tabName === 'payments') {
                document.getElementById('paymentsTab').classList.add('active');
                document.querySelector('.tab:nth-child(1)').classList.add('active');
            } else if (tabName === 'iban') {
                document.getElementById('ibanTab').classList.add('active');
                document.querySelector('.tab:nth-child(2)').classList.add('active');
            }
        }
