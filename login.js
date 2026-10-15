const loginForm = document.getElementById("loginForm");
const loginButton = document.getElementById("loginButton");
const loginMessage = document.getElementById("loginMessage");


if (loginForm) {

    loginForm.addEventListener("submit", async (event) => {

        event.preventDefault();

        const email =
            document.getElementById("email").value.trim();

        const password =
            document.getElementById("password").value;


        if (!email || !password) {

            showLoginMessage(
                "Email dan password wajib diisi.",
                "error"
            );

            return;
        }

        try {

            loginButton.disabled = true;

            loginButton.innerHTML = `
                <i class="fa-solid fa-spinner fa-spin"></i>
                <span>Memproses...</span>
            `;

            showLoginMessage("", "");

            const {
                data,
                error
            } = await supabaseClient.auth.signInWithPassword({
                email,
                password
            });


            if (error) {

                console.error("LOGIN ERROR:", error);

                showLoginMessage(
                    error.message,
                    "error"
                );

                return;
            }


            console.log(
                "Login berhasil:",
                data
            );


            showLoginMessage(
                "Login berhasil. Mengarahkan ke dashboard...",
                "success"
            );


            setTimeout(() => {

                window.location.href =
                    "./index.html";

            }, 500);


        } catch (error) {

            console.error(
                "LOGIN EXCEPTION:",
                error
            );

            showLoginMessage(
                "Terjadi kesalahan saat login.",
                "error"
            );


        } finally {

            loginButton.disabled = false;

            loginButton.innerHTML = `
                <span>Masuk</span>
                <i class="fa-solid fa-arrow-right"></i>
            `;

        }

    });

}

const togglePassword =
    document.getElementById("togglePassword");

const passwordInput =
    document.getElementById("password");


if (
    togglePassword &&
    passwordInput
) {

    togglePassword.addEventListener(
        "click",
        () => {

            const isPassword =
                passwordInput.type === "password";


            passwordInput.type =
                isPassword
                    ? "text"
                    : "password";


            togglePassword.innerHTML =
                isPassword
                    ? '<i class="fa-regular fa-eye-slash"></i>'
                    : '<i class="fa-regular fa-eye"></i>';

        }
    );

}

function showLoginMessage(
    message,
    type
) {

    if (!loginMessage) return;


    loginMessage.textContent =
        message;


    loginMessage.style.color =
        type === "error"
            ? "#dc3545"
            : type === "success"
                ? "#198754"
                : "";

}