        // AUTHENTICATION
        const CREDENTIALS = {
            username: 'AuxiliumUser',
            password: 'Auxilium2026!'
        };

        document.getElementById('loginForm').addEventListener('submit', function (e) {
            e.preventDefault();

            const username = document.getElementById('username').value;
            const password = document.getElementById('password').value;
            const errorMessage = document.getElementById('loginError');

            if (username === CREDENTIALS.username && password === CREDENTIALS.password) {
                errorMessage.classList.remove('show');
                document.getElementById('loginContainer').style.display = 'none';
                document.getElementById('mainContainer').classList.add('show');
                sessionStorage.setItem('authenticated', 'true');
            } else {
                errorMessage.classList.add('show');
                document.getElementById('username').classList.add('error');
                document.getElementById('password').classList.add('error');

                setTimeout(() => {
                    document.getElementById('username').classList.remove('error');
                    document.getElementById('password').classList.remove('error');
                }, 2000);
            }
        });

        function logout() {
            sessionStorage.removeItem('authenticated');
            document.getElementById('mainContainer').classList.remove('show');
            document.getElementById('loginContainer').style.display = 'block';
            document.getElementById('username').value = '';
            document.getElementById('password').value = '';
            document.getElementById('loginError').classList.remove('show');

            // Reset tabs
            switchTab('payments');
        }

        window.addEventListener('load', function () {
            if (sessionStorage.getItem('authenticated') === 'true') {
                document.getElementById('loginContainer').style.display = 'none';
                document.getElementById('mainContainer').classList.add('show');
            }
        });
