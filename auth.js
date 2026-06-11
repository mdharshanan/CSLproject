(function () {
    const loginKey = "lmsLoggedIn";
    const roleKey = "lmsRole";
    const adminEmail = "admin@lms.edu";
    const adminPassword = "Admin@123";
    const currentScript = document.currentScript;
    const isProtectedPage = currentScript && currentScript.dataset.protected === "true";
    const loginForm = document.querySelector(".login-card");
    const registerForm = document.getElementById("register-form");
    const message = document.querySelector("[data-login-message]");
    const registerMessage = document.querySelector("[data-register-message]");
    const params = new URLSearchParams(window.location.search);
    const logoutButtons = document.querySelectorAll("[data-logout]");

    function isLoggedIn() {
        return localStorage.getItem(loginKey) === "true";
    }

    function getRole() {
        return localStorage.getItem(roleKey) || "guest";
    }

    function isAdmin() {
        return getRole() === "admin";
    }

    function setLogin(role) {
        localStorage.setItem(loginKey, "true");
        localStorage.setItem(roleKey, role);
    }

    function currentPage() {
        return window.location.pathname.split("/").pop() || "index.html";
    }

    function redirectToLogin(role) {
        const target = currentPage() + window.location.search + window.location.hash;
        const roleParam = role === 'admin' ? '&role=admin' : '';
        window.location.replace("login.html?next=" + encodeURIComponent(target) + roleParam);
    }

    function updateLogoutButtons() {
        logoutButtons.forEach(function (button) {
            button.style.display = isLoggedIn() ? "inline-flex" : "none";
        });
    }

    function updateAdminLinks() {
        document.querySelectorAll('.admin-link').forEach(function (link) {
            link.style.display = isAdmin() ? "inline-flex" : "none";
        });
    }

    function updateLoginState() {
        // Toggle body class so CSS rules handle login-nav-btn vs logout-button
        if (isLoggedIn()) {
            document.body.classList.add('user-logged-in');
        } else {
            document.body.classList.remove('user-logged-in');
        }
        updateLogoutButtons();
        updateAdminLinks();
    }

    if (currentPage() === "index.html") {
        document.body.classList.add("index-page");
    }

    updateLoginState();

    if (currentPage() === "admin.html" && !isAdmin()) {
        redirectToLogin('admin');
        return;
    }

    if (isProtectedPage && !isLoggedIn()) {
        redirectToLogin();
        return;
    }

    if (loginForm) {
        loginForm.addEventListener("submit", async function (event) {
            event.preventDefault();
            if (message) {
                message.textContent = '';
            }

            const email = loginForm.querySelector("#email");
            const password = loginForm.querySelector("#password");
            const emailValue = email.value.trim().toLowerCase();
            const passwordValue = password.value.trim();

            if (!emailValue || !passwordValue) {
                if (message) {
                    message.textContent = "Please enter email and password.";
                }
                return;
            }

            try {
                const response = await fetch('/api/login', {
                    method: 'POST',
                    credentials: 'same-origin',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ email: emailValue, password: passwordValue }),
                });

                const result = await response.json();
                if (!response.ok) {
                    message.textContent = result.error || 'Login failed. Please check your credentials.';
                    return;
                }

                setLogin(result.role || 'user');
                const nextPage = params.get("next") || (result.role === 'admin' ? "admin.html" : "courses.html");
                window.location.href = nextPage;
            } catch (error) {
                console.error(error);
                message.textContent = 'Unable to reach the server. Please start the backend and try again.';
            }
        });
    }

        if (registerForm) {
            registerForm.addEventListener("submit", async function (event) {
                event.preventDefault();
                if (registerMessage) {
                    registerMessage.textContent = '';
                }

                const name = registerForm.querySelector("#name");
                const email = registerForm.querySelector("#email");
                const mobile = registerForm.querySelector("#mobile");
                const password = registerForm.querySelector("#password");
                const nameValue = name.value.trim();
                const emailValue = email.value.trim().toLowerCase();
                const mobileValue = mobile ? mobile.value.trim() : "";
                const passwordValue = password.value.trim();

                if (!nameValue || !emailValue || !passwordValue) {
                    if (registerMessage) {
                        registerMessage.textContent = "Please fill all fields (mobile is optional).";
                    }
                    return;
                }

                try {
                    const response = await fetch('/api/register', {
                        method: 'POST',
                        credentials: 'same-origin',
                        headers: {
                            'Content-Type': 'application/json',
                        },
                        body: JSON.stringify({ name: nameValue, email: emailValue, password: passwordValue }),
                    });

                    const result = await response.json();
                    if (!response.ok) {
                        registerMessage.textContent = result.error || 'Registration failed. Please check your details.';
                        return;
                    }

                    setLogin(result.role || 'user');
                    const nextPage = params.get("next") || (result.role === 'admin' ? "admin.html" : "courses.html");
                    window.location.href = nextPage;
                } catch (error) {
                    console.error(error);
                    if (registerMessage) {
                        registerMessage.textContent = 'Unable to reach the server. Please start the backend and try again.';
                    }
                }
            });
        }

        updateLoginState();

        document.querySelectorAll("[data-logout]").forEach(function (button) {
            button.addEventListener("click", async function (event) {
                event.preventDefault();
                try {
                    await fetch('/api/logout', {
                        method: 'POST',
                        credentials: 'same-origin',
                    });
                } catch (e) {
                    console.warn('Failed to notify server logout:', e);
                }
                localStorage.removeItem(loginKey);
                localStorage.removeItem(roleKey);
                updateLoginState();
                window.location.href = "index.html";
            });
        });
    })();
