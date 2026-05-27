document.addEventListener('DOMContentLoaded', () => {
  const loginForm = document.getElementById('login-form');
  const togglePasswordBtn = document.getElementById('toggle-password-btn');
  const passwordInput = document.getElementById('password');
  const usernameInput = document.getElementById('username');
  const errorMsg = document.getElementById('login-error-msg');
  const submitBtn = document.getElementById('login-submit-btn');

  // Pre-fill username if remembered
  const rememberedUser = localStorage.getItem('skilltrack_remembered_user');
  if (rememberedUser && usernameInput) {
    usernameInput.value = rememberedUser;
    const rememberMeCheckbox = document.getElementById('remember-me');
    if (rememberMeCheckbox) rememberMeCheckbox.checked = true;
  }

  // Toggle password visibility
  if (togglePasswordBtn && passwordInput) {
    togglePasswordBtn.addEventListener('click', () => {
      const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
      passwordInput.setAttribute('type', type);
      
      const svg = togglePasswordBtn.querySelector('.eye-icon');
      if (svg) {
        svg.style.opacity = type === 'text' ? '0.5' : '1';
      }
    });
  }

  // Handle submit
  if (loginForm) {
    loginForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      
      const usernameVal = usernameInput.value.trim();
      const passwordVal = passwordInput.value;
      
      // Clear error states
      errorMsg.classList.add('hidden');
      errorMsg.innerText = '';
      
      const originalBtnText = submitBtn.innerHTML;
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span>Verifying...</span>';
      
      try {
        const response = await fetch('/api/login', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            username: usernameVal,
            password: passwordVal
          })
        });

        const data = await response.json();
        
        if (response.ok && data.success) {
          // Handle "Remember me"
          const rememberMe = document.getElementById('remember-me').checked;
          if (rememberMe) {
            localStorage.setItem('skilltrack_remembered_user', usernameVal);
          } else {
            localStorage.removeItem('skilltrack_remembered_user');
          }
          
          // Redirect to index page
          window.location.href = '/';
        } else {
          // Show error message
          errorMsg.classList.remove('hidden');
          errorMsg.innerText = data.message || 'Invalid username or password.';
          
          // Clear password input and focus
          passwordInput.value = '';
          passwordInput.focus();
          
          // Shake the card to indicate validation failure
          const card = document.querySelector('.login-card');
          if (card) {
            card.style.animation = 'none';
            void card.offsetWidth; // Trigger reflow
            card.style.animation = 'shake 0.4s ease-in-out';
          }
        }
      } catch (err) {
        console.error('Authentication error:', err);
        errorMsg.classList.remove('hidden');
        errorMsg.innerText = 'Unable to connect to the authentication server. Please try again.';
      } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML = originalBtnText;
      }
    });
  }
});
