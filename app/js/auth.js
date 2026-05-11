/* ============================================================
   AUTH.JS - Autenticacion y gestion de negocio
   ============================================================ */

const Auth = (() => {
    let currentUser = null;

    function init() {
        setupUI();
        setupAuthObserver();
    }

    // --- UI Event Listeners ---
    function setupUI() {
        // Toggle login <-> register
        document.getElementById('link-register').addEventListener('click', e => {
            e.preventDefault();
            document.getElementById('auth-login').style.display = 'none';
            document.getElementById('auth-register').style.display = 'block';
            clearErrors();
        });

        document.getElementById('link-login').addEventListener('click', e => {
            e.preventDefault();
            document.getElementById('auth-login').style.display = 'block';
            document.getElementById('auth-register').style.display = 'none';
            clearErrors();
        });

        // Login
        document.getElementById('btn-login').addEventListener('click', login);
        document.getElementById('login-password').addEventListener('keydown', e => {
            if (e.key === 'Enter') login();
        });

        // Register
        document.getElementById('btn-register').addEventListener('click', register);
        document.getElementById('register-password2').addEventListener('keydown', e => {
            if (e.key === 'Enter') register();
        });

        // Business setup
        document.getElementById('btn-create-business').addEventListener('click', createBusiness);
        document.getElementById('btn-join-business').addEventListener('click', joinBusiness);
        document.getElementById('join-code').addEventListener('keydown', e => {
            if (e.key === 'Enter') joinBusiness();
        });

        document.getElementById('link-logout-setup').addEventListener('click', e => {
            e.preventDefault();
            auth.signOut();
        });

        // Logout
        document.getElementById('btn-logout').addEventListener('click', () => {
            auth.signOut();
        });
    }

    // --- Auth State Observer ---
    function setupAuthObserver() {
        auth.onAuthStateChanged(async (user) => {
            if (user) {
                currentUser = user;
                try {
                    const userDoc = await db.collection('usuarios').doc(user.uid).get();
                    if (userDoc.exists && userDoc.data().negocioId) {
                        enterApp(user, userDoc.data().negocioId);
                    } else {
                        showBusinessSetup();
                    }
                } catch (err) {
                    console.error('Error verificando usuario:', err);
                    if (err.code === 'permission-denied') {
                        showError('auth-error', 'Error de permisos en la base de datos. Contacta al administrador.');
                    } else {
                        showError('auth-error', 'Error de conexion. Recarga la pagina.');
                    }
                }
            } else {
                currentUser = null;
                exitApp();
                showLogin();
            }
        });
    }

    // --- Screen switching ---
    function showLogin() {
        document.getElementById('auth-screen').style.display = 'flex';
        document.getElementById('business-screen').style.display = 'none';
        document.getElementById('app-container').style.display = 'none';
    }

    function showBusinessSetup() {
        document.getElementById('auth-screen').style.display = 'none';
        document.getElementById('business-screen').style.display = 'flex';
        document.getElementById('app-container').style.display = 'none';
    }

    function enterApp(user, negocioId) {
        document.getElementById('auth-screen').style.display = 'none';
        document.getElementById('business-screen').style.display = 'none';
        document.getElementById('app-container').style.display = 'flex';

        // Update sidebar user info
        document.getElementById('user-email').textContent = user.email;
        document.getElementById('negocio-code-display').textContent = negocioId;

        // Connect data layer to Firestore
        DataStore.connect(negocioId);
    }

    function exitApp() {
        DataStore.disconnect();
    }

    // --- Login ---
    async function login() {
        const email = document.getElementById('login-email').value.trim();
        const password = document.getElementById('login-password').value;

        if (!email || !password) {
            showError('auth-error', 'Completa todos los campos');
            return;
        }

        try {
            setLoading('btn-login', true);
            await auth.signInWithEmailAndPassword(email, password);
        } catch (err) {
            showError('auth-error', translateError(err.code));
        } finally {
            setLoading('btn-login', false);
        }
    }

    // --- Register ---
    async function register() {
        const email = document.getElementById('register-email').value.trim();
        const password = document.getElementById('register-password').value;
        const password2 = document.getElementById('register-password2').value;

        if (!email || !password || !password2) {
            showError('auth-error', 'Completa todos los campos');
            return;
        }

        if (password !== password2) {
            showError('auth-error', 'Las contrasenas no coinciden');
            return;
        }

        if (password.length < 6) {
            showError('auth-error', 'La contrasena debe tener al menos 6 caracteres');
            return;
        }

        try {
            setLoading('btn-register', true);
            await auth.createUserWithEmailAndPassword(email, password);
        } catch (err) {
            showError('auth-error', translateError(err.code));
        } finally {
            setLoading('btn-register', false);
        }
    }

    // --- Business creation ---
    function generateCode() {
        const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
        let code = '';
        for (let i = 0; i < 6; i++) {
            code += chars.charAt(Math.floor(Math.random() * chars.length));
        }
        return code;
    }

    async function createBusiness() {
        const code = generateCode();
        const user = auth.currentUser;

        try {
            setLoading('btn-create-business', true);

            // Check for existing localStorage data to migrate
            let initialData = DataStore.getDefaults();
            const existing = localStorage.getItem('motoflip_data');
            if (existing) {
                try {
                    const parsed = JSON.parse(existing);
                    initialData = {
                        ...initialData,
                        ...parsed,
                        config: { ...initialData.config, ...(parsed.config || {}) }
                    };
                } catch {}
            }

            // Create business document
            await db.collection('negocios').doc(code).set({
                config: initialData.config,
                motos: initialData.motos || [],
                gastos: initialData.gastos || [],
                nextMotoId: initialData.nextMotoId || 1,
                nextGastoId: initialData.nextGastoId || 1,
                members: [user.uid],
                createdBy: user.uid,
                createdAt: firebase.firestore.FieldValue.serverTimestamp()
            });

            // Link user to business
            await db.collection('usuarios').doc(user.uid).set({
                email: user.email,
                negocioId: code
            });

            enterApp(user, code);

            // Show code to share
            setTimeout(() => {
                alert(
                    'Negocio creado!\n\n' +
                    'Tu codigo para compartir es:\n\n' +
                    '   ' + code + '\n\n' +
                    'Pasale este codigo a tu socio para que se una.'
                );
            }, 600);

        } catch (err) {
            showError('business-error', 'Error al crear negocio: ' + err.message);
        } finally {
            setLoading('btn-create-business', false);
        }
    }

    // --- Join business ---
    async function joinBusiness() {
        const code = document.getElementById('join-code').value.trim().toUpperCase();
        const user = auth.currentUser;

        if (!code || code.length < 4) {
            showError('business-error', 'Ingresa un codigo valido');
            return;
        }

        try {
            setLoading('btn-join-business', true);

            const negocioDoc = await db.collection('negocios').doc(code).get();

            if (!negocioDoc.exists) {
                showError('business-error', 'No se encontro un negocio con ese codigo');
                return;
            }

            // Add user as member
            const members = negocioDoc.data().members || [];
            if (!members.includes(user.uid)) {
                members.push(user.uid);
                await db.collection('negocios').doc(code).update({ members });
            }

            // Link user to business
            await db.collection('usuarios').doc(user.uid).set({
                email: user.email,
                negocioId: code
            });

            enterApp(user, code);

        } catch (err) {
            showError('business-error', 'Error al unirse: ' + err.message);
        } finally {
            setLoading('btn-join-business', false);
        }
    }

    // --- Helpers ---
    function showError(elementId, message) {
        const el = document.getElementById(elementId);
        el.textContent = message;
        el.style.display = 'block';
        setTimeout(() => { el.style.display = 'none'; }, 5000);
    }

    function clearErrors() {
        document.querySelectorAll('.auth-error').forEach(el => {
            el.style.display = 'none';
        });
    }

    function setLoading(btnId, loading) {
        const btn = document.getElementById(btnId);
        if (loading) {
            btn.dataset.originalText = btn.textContent;
            btn.textContent = 'Cargando...';
            btn.disabled = true;
        } else {
            btn.textContent = btn.dataset.originalText || btn.textContent;
            btn.disabled = false;
        }
    }

    function translateError(code) {
        const map = {
            'auth/email-already-in-use': 'Ese email ya esta registrado',
            'auth/invalid-email': 'Email invalido',
            'auth/weak-password': 'La contrasena es muy debil (minimo 6 caracteres)',
            'auth/user-not-found': 'No existe una cuenta con ese email',
            'auth/wrong-password': 'Contrasena incorrecta',
            'auth/too-many-requests': 'Demasiados intentos. Espera un momento',
            'auth/invalid-credential': 'Email o contrasena incorrectos'
        };
        return map[code] || 'Error: ' + code;
    }

    return { init };
})();
