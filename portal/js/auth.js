const API_URL = 'https://digitcom-backend.onrender.com/api';
window.API_URL = API_URL; // Exposure for other scripts

const loginForm = document.getElementById('login-form');
const errorMsg = document.getElementById('error-msg');

if (loginForm) {
    loginForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const username = document.getElementById('username').value;
        const password = document.getElementById('password').value;
        const loginBtn = document.getElementById('login-btn');
        const originalText = loginBtn ? loginBtn.textContent : 'Sign In';
        
        if (loginBtn) {
            loginBtn.disabled = true;
            loginBtn.textContent = 'Signing in... (Render may take 30-50s to wake up)';
        }
        errorMsg.style.display = 'none';
        
        console.log(`Connecting to: ${API_URL}/auth/login`);
        
        try {
            const response = await fetch(`${API_URL}/auth/login`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ username, password })
            });
            
            console.log('Response Status:', response.status);
            const data = await response.json();
            console.log('Response Data:', data);
            
            if (response.ok) {
                // Save token and user info
                localStorage.setItem('token', data.token);
                localStorage.setItem('user', JSON.stringify(data.user));
                console.log('Login successful, redirecting...');
                
                // Redirect based on role
                if (data.user.role === 'ADMIN') {
                    window.location.href = 'dashboard.html';
                } else {
                    window.location.href = 'warehouse.html';
                }
            } else {
                errorMsg.style.display = 'block';
                errorMsg.textContent = data.error || 'Login failed.';
                if (loginBtn) {
                    loginBtn.disabled = false;
                    loginBtn.textContent = originalText;
                }
            }
        } catch (error) {
            errorMsg.style.display = 'block';
            errorMsg.textContent = 'Server connection failed. (Check your internet or if the server is awake)';
            console.error('Login Error:', error);
            if (loginBtn) {
                loginBtn.disabled = false;
                loginBtn.textContent = originalText;
            }
        }
    });
}

// Global check for logged-in status
function checkAuth() {
    const token = localStorage.getItem('token');
    if (!token && !window.location.pathname.endsWith('index.html')) {
        window.location.href = 'index.html';
    }
}

function logout() {
    localStorage.removeItem('token');
    localStorage.removeItem('user');
    window.location.href = 'index.html';
}
