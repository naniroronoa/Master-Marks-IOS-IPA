// --- نظام التفعيل والحماية V2 (Activation & License System) ---
// SECRET KEY for simple XOR (must match Keygen)
const LICENSE_SECRET = "TEACHER_XP_2026_SECRET";

let activationState = {
    isActivated: false,
    hwid: '',
    plan: '',
    expiry: 0,
    activationKey: ''
};

async function initActivation() {
    console.log("Activation: Initializing V2...");
    const hwidDisplay = document.getElementById('hwid-display');

    try {
        if (window.electronAPI && window.electronAPI.getHWID) {
            activationState.hwid = await window.electronAPI.getHWID();
        } else {
            activationState.hwid = getFallbackHWID();
        }
    } catch (error) {
        console.error("Activation Error:", error);
        activationState.hwid = getFallbackHWID();
    }

    if (hwidDisplay) hwidDisplay.value = activationState.hwid;

    // التحقق من المفتاح المحفوظ
    const savedKey = localStorage.getItem('teacher_activation_key');
    if (savedKey) {
        const licenseData = verifyAndDecryptKey(savedKey, activationState.hwid);
        if (licenseData && !isExpired(licenseData.expiry)) {
            activationState.isActivated = true;
            activationState.plan = licenseData.plan;
            activationState.expiry = licenseData.expiry;
            activationState.activationKey = savedKey;
            document.body.classList.add('activated');

            // إظهار زر معلومات الترخيص وإخفاء زر التفعيل
            updateLicenseUI(true);
            updateSplashUI(true); // Splash: Show Login Button

            if (typeof window.switchSection === 'function' && !document.body.classList.contains('home-init-done')) {
                window.switchSection(null, 'home');
                document.body.classList.add('home-init-done');
            }

            // إعداد التنبيه للإغلاق التلقائي عند انتهاء الصلاحية
            setupExpirationTimer(activationState.expiry);
        } else {
            // منتهي الصلاحية أو غير صالح - تحقق مما إذا كان يجب إظهار رسالة الشكر
            if (savedKey && licenseData && isExpired(licenseData.expiry)) {
                checkAndShowExpirationNotice(licenseData);
            }
            handleInvalidLicense();
        }
    } else {
        handleInvalidLicense();
    }

    // مستمع للتنبيه
    // مستمع للتنبيه (Updated Logic)
    document.addEventListener('click', (e) => {
        // If activated, do nothing
        if (activationState.isActivated) return;

        // If passed activation check in other ways
        if (document.body.classList.contains('activated')) return;

        // If clicking inside activation UI, do nothing
        const isActivationUI = e.target.closest('#activation-section') ||
            e.target.closest('#activation-modal') ||
            e.target.closest('#nav-activation') ||
            e.target.closest('.modal-overlay.open'); // Don't show if a modal is open

        if (!isActivationUI) {
            showActivationToast('يرجى تفعيل البرنامج للوصول إلى كافة الميزات', 'error');
        }
    });

    // تحديث العد التنازلي إذا كانت النافذة مفتوحة
    setInterval(updateCountdownDisplay, 60000); // كل دقيقة
}

function handleInvalidLicense() {
    activationState.isActivated = false;
    document.body.classList.remove('activated');
    updateLicenseUI(false);
    updateSplashUI(false); // Splash: Show Activate Button
    localStorage.removeItem('teacher_activation_key');
    if (typeof window.switchSection === 'function') {
        window.switchSection(null, 'activation');
    }
}

function updateLicenseUI(isActivated) {
    const navActivation = document.getElementById('nav-activation');
    const headerActionBtns = document.querySelector('.header-action-btns');
    const appLogoContent = document.getElementById('app-logo-content');

    if (isActivated) {
        if (navActivation) navActivation.style.display = 'none';
        if (headerActionBtns) headerActionBtns.style.display = 'flex';
        if (appLogoContent) appLogoContent.style.display = 'none'; // Hide default logo
    } else {
        if (navActivation) navActivation.style.display = 'flex';
        if (headerActionBtns) headerActionBtns.style.display = 'none';
        if (appLogoContent) appLogoContent.style.display = 'flex'; // Show default logo
    }
}

// فك تشفير المفتاح والتحقق منه
function verifyAndDecryptKey(key, currentHwid) {
    if (!key.startsWith('TLILI - ')) return null;

    try {
        const body = key.substring(8); // Remove "TLILI - "
        const targetAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}|;:,.<>?";
        const b64Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

        // 1. Map back to Base64
        let b64 = "";
        for (let i = 0; i < body.length; i++) {
            const index = targetAlphabet.indexOf(body[i]);
            // If the index is within B64 range, it's part of our data
            // (Padding characters might have higher indices but shouldn't interfere if we split correctly)
            if (index >= 0 && index < 64) {
                b64 += b64Alphabet[index];
            } else {
                // If it's a padding character (index >= 64), we might have reached the end of encoded data
                // However, our packed data usually ends before the 50 char limit.
                // Decrypting more doesn't hurt as long as the split('|') works.
                break;
            }
        }

        // 2. Decrypt
        // We don't know the salt yet, but it's the first part of the decrypted string
        // This is tricky. Let's assume a fixed LICENSE_SECRET for the first XOR 
        // Or wait, my keygen used `LICENSE_SECRET + salt` for XOR.
        // To decrypt, we need the salt.

        // RE-EVALUATION: The salt must be sent in a way we can see it before XORing everything with it.
        // Better approach: XOR everything with LICENSE_SECRET alone, or send salt in clear.
        // Let's use XOR with LICENSE_SECRET only for the metadata, but include salt in the metadata.

        const decrypted = xorDecrypt(atob(b64), LICENSE_SECRET);
        const parts = decrypted.split('|');
        if (parts.length < 4) return null;

        const [salt, hwidDigits, planChar, expiry] = parts;

        // 3. Verify HWID
        const currentHwidDigits = currentHwid.replace(/\D/g, '');
        if (hwidDigits !== currentHwidDigits) return null;

        return {
            plan: planChar === 'P' ? 'PREMIUM' : 'BASIC',
            expiry: parseInt(expiry)
        };
    } catch (e) {
        console.error("Activation: Decryption failed", e);
        return null;
    }
}

function xorDecrypt(str, key) {
    let result = '';
    for (let i = 0; i < str.length; i++) {
        result += String.fromCharCode(str.charCodeAt(i) ^ key.charCodeAt(i % key.length));
    }
    return result;
}

function isExpired(expiryTimestamp) {
    return Date.now() > expiryTimestamp;
}

function getFallbackHWID() {
    let id = localStorage.getItem('teacher_fallback_hwid');

    // التنسيق الجديد: Txxx - Lxxx - Ixxx - Lxxx - Ixxx
    const isValidNewFormat = id && /^T\d{3} - L\d{3} - I\d{3} - L\d{3} - I\d{3}$/.test(id);

    if (id && !isValidNewFormat) {
        localStorage.removeItem('teacher_fallback_hwid');
        id = null;
    }

    if (!id) {
        // إنشاء معرف طوارئ بالتنسيق المعتمد
        const rand = () => Math.floor(100 + Math.random() * 900);
        id = `T${rand()} - L${rand()} - I${rand()} - L${rand()} - I${rand()}`;
        localStorage.setItem('teacher_fallback_hwid', id);
    }
    return id;
}

// واجهة معلومات الترخيص
window.showLicenseInfo = function () {
    if (!activationState.isActivated) return;

    document.getElementById('license-info-modal').classList.add('open');
    refreshLicenseInfo();
};

window.showActivationModal = function () {
    const modal = document.getElementById('activation-modal');
    if (modal) modal.classList.add('open');
};

window.closeActivationModal = function () {
    const modal = document.getElementById('activation-modal');
    if (modal) modal.classList.remove('open');
};

window.closeLicenseModal = function () {
    const modal = document.getElementById('license-info-modal');
    if (modal) modal.classList.remove('open');
};

window.refreshLicenseInfo = function () {
    const planText = activationState.plan === 'PREMIUM' ? 'ممتاز (Premium)' : 'عادي (Basic)';
    document.getElementById('license-plan').innerText = planText;
    document.getElementById('license-plan').className = 'value plan-badge ' + activationState.plan.toLowerCase();

    const expiryDate = new Date(activationState.expiry);

    // Format Date: DD/MM/YYYY
    const dateOptions = { year: 'numeric', month: '2-digit', day: '2-digit' };
    document.getElementById('license-expiry').innerText = expiryDate.toLocaleDateString('ar-DZ', dateOptions);

    // Format Time: HH:mm:ss (24h)
    const timeOptions = { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false };
    const timeString = expiryDate.toLocaleTimeString('en-US', timeOptions); // en-US to ensure 24h digits are standard

    const timeEl = document.getElementById('license-expiry-time');
    if (timeEl) timeEl.innerText = timeString;

    document.getElementById('license-hwid').innerText = activationState.hwid;

    updateCountdownDisplay();
};

function updateCountdownDisplay() {
    const now = Date.now();
    let diff = activationState.expiry - now;

    if (diff <= 0) {
        // التحقق من إظهار رسالة الشكر فور الانتهاء والبرنامج مفتوح
        checkAndShowExpirationNotice({
            plan: activationState.plan,
            expiry: activationState.expiry
        });
        handleInvalidLicense();
        closeLicenseModal();

        // إغلاق التطبيق تلقائياً عند انتهاء الصلاحية
        return;
    }

    const minutes = Math.floor((diff / (1000 * 60)) % 60);
    const hours = Math.floor((diff / (1000 * 60 * 60)) % 24);
    const days = Math.floor((diff / (1000 * 60 * 60 * 24)) % 30);
    const months = Math.floor(diff / (1000 * 60 * 60 * 24 * 30));

    const cdMonths = document.getElementById('cd-months');
    const cdDays = document.getElementById('cd-days');
    const cdHours = document.getElementById('cd-hours');
    const cdMins = document.getElementById('cd-mins');

    if (cdMonths) cdMonths.innerText = months;
    if (cdDays) cdDays.innerText = days;
    if (cdHours) cdHours.innerText = hours;
    if (cdMins) cdMins.innerText = minutes;
}

window.confirmCancelLicense = function () {
    const modal = document.getElementById('license-cancel-modal');
    if (modal) modal.classList.add('open');
};

window.closeCancelLicenseModal = function () {
    const modal = document.getElementById('license-cancel-modal');
    if (modal) modal.classList.remove('open');
};

window.executeCancelLicense = function () {
    handleInvalidLicense();
    closeCancelLicenseModal();
    closeLicenseModal();
    // Optional: Show a small toast or alert confirming cancellation
    // alert('تم إلغاء الترخيص بنجاح.');
};

window.verifyActivation = function () {
    const inputKey = document.getElementById('activation-key-input').value.trim();
    if (!inputKey) return;

    const licenseData = verifyAndDecryptKey(inputKey, activationState.hwid);
    if (licenseData) {
        if (isExpired(licenseData.expiry)) {
            window.showActivationErrorModal('EXPIRED');
            return;
        }

        localStorage.setItem('teacher_activation_key', inputKey);

        // حفظ بيانات التفعيل الأخير لإظهارها عند الانتهاء
        const lastActivation = {
            plan: licenseData.plan,
            activatedDate: Date.now(),
            expiry: licenseData.expiry,
            notified: false
        };
        localStorage.setItem('last_activation_info', JSON.stringify(lastActivation));

        activationState.isActivated = true;
        activationState.plan = licenseData.plan;
        activationState.expiry = licenseData.expiry;
        activationState.activationKey = inputKey;

        document.body.classList.add('activated');
        updateLicenseUI(true);
        window.closeActivationModal();

        // إعداد التنبيه للإغلاق التلقائي عند انتهاء الصلاحية للمفتاح الجديد
        setupExpirationTimer(activationState.expiry);

        // Show Thank You Modal instead of Alert
        window.showThankYouModal(licenseData.plan, licenseData.expiry);

    } else {
        console.log("Invalid key detected");
        window.showActivationErrorModal('INVALID');
    }
};

window.showActivationErrorModal = function (type) {
    const modal = document.getElementById('activation-error-modal');
    if (!modal) return;

    const titleEl = document.getElementById('error-modal-title');
    const msgEl = document.getElementById('error-modal-message');

    if (type === 'EXPIRED') {
        if (titleEl) titleEl.textContent = 'المفتاح منتهي الصلاحية';
        if (msgEl) msgEl.textContent = 'عذراً، هذا المفتاح منتهي الصلاحية. يرجى تجديد اشتراكك للحصول على مفتاح جديد.';
    } else {
        if (titleEl) titleEl.textContent = 'كود التفعيل خاطئ';
        if (msgEl) msgEl.textContent = 'كود التفعيل غير صحيح أو لا يتوافق مع هذا الجهاز. يرجى التأكد من الكود والمحاولة مرة أخرى.';
    }

    modal.classList.add('open');
};

window.closeActivationErrorModal = function () {
    const modal = document.getElementById('activation-error-modal');
    if (modal) modal.classList.remove('open');
};

window.showThankYouModal = function (plan, expiry) {
    const modal = document.getElementById('thank-you-modal');
    if (!modal) return;

    // Populate data
    const planText = plan === 'PREMIUM' ? 'نسخة كاملة (Premium)' : 'نسخة أساسية (Basic)';
    const expiryDate = new Date(expiry).toLocaleDateString('ar-DZ', { year: 'numeric', month: 'long', day: 'numeric' });

    const elPlan = document.getElementById('ty-license-plan');
    const elExpiry = document.getElementById('ty-license-expiry');

    if (elPlan) elPlan.textContent = planText;
    if (elExpiry) elExpiry.textContent = expiryDate;

    // Show modal
    modal.classList.add('open');
};

window.closeThankYouModal = function () {
    const modal = document.getElementById('thank-you-modal');
    if (modal) modal.classList.remove('open');

    // Proceed to home section logic
    if (typeof window.switchSection === 'function') {
        window.switchSection(null, 'home');
    }
};

window.copyHWID = function () {
    const hwid = activationState.hwid || "جاري التحميل...";
    navigator.clipboard.writeText(hwid).then(() => {
        // Show rich modal instead of alert/toast
        showHWIDCopyModal();
    }).catch(err => {
        console.error('Failed to copy: ', err);
        alert('فشل نسخ المعرف. يرجى نسخه يدوياً.');
    });
};

window.showHWIDCopyModal = function () {
    const modal = document.getElementById('hwid-copy-modal');
    if (modal) {
        modal.classList.add('open');
        // Optional: Play success sound if available
        // if (typeof playSuccessSound === 'function') playSuccessSound();
    }
};

window.closeHWIDCopyModal = function () {
    const modal = document.getElementById('hwid-copy-modal');
    if (modal) modal.classList.remove('open');
};

window.refreshHWID = async function () {
    const hwidDisplay = document.getElementById('hwid-display');
    if (hwidDisplay) hwidDisplay.value = "جاري التحديث...";
    await initActivation();
};

window.showActivationToast = function (message, type = 'error') {
    const toast = document.getElementById('activation-toast');
    if (!toast) return;

    // Reset classes
    toast.className = 'activation-toast';

    // Set content and type
    toast.textContent = message || 'يرجى تفعيل البرنامج';

    if (type === 'success') {
        toast.classList.add('success');
    }

    toast.classList.add('show');

    // Hide after 3 seconds
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
};

// --- نظام الإغلاق التلقائي ---
let expirationTimeout = null;

function setupExpirationTimer(expiryTimestamp) {
    if (expirationTimeout) clearTimeout(expirationTimeout);

    const now = Date.now();
    const timeLeft = expiryTimestamp - now;

    // JavaScript setTimeout has a maximum delay of 2,147,483,647ms (~24.8 days)
    // If the delay is larger, it's treated as 0 or triggers immediately.
    const MAX_TIMEOUT = 2147483647;

    if (timeLeft <= 0) {
        // Already expired, handled by init or updateCountdownDisplay
        return;
    }

    if (timeLeft > MAX_TIMEOUT) {
        // If the expiration is too far in the future (> 24.8 days), we don't set a timeout.
        // The minute-by-minute interval (updateCountdownDisplay) in initActivation()
        // will safely handle the expiration when the time comes.
        console.log("Activation: Expiry is in more than 24 days. Relying on periodic check.");
        return;
    }

    // Set auto-close timer only if within safe bounds
    console.log(`Activation: Setting auto-close timer for ${Math.round(timeLeft / 1000)} seconds.`);
    
    expirationTimeout = setTimeout(() => {
        console.log("Activation: License expired exactly now. Triggering auto-close...");
        
        checkAndShowExpirationNotice({
            plan: activationState.plan,
            expiry: activationState.expiry
        });
        
        handleInvalidLicense();
        closeLicenseModal();

    }, timeLeft);
}


// --- منطق شكر المستخدم عند انتهاء الصلاحية ---
function checkAndShowExpirationNotice(licenseData) {
    if (!licenseData) return;

    const lastInfoStr = localStorage.getItem('last_activation_info');
    let lastInfo = lastInfoStr ? JSON.parse(lastInfoStr) : null;

    // إذا لم تتوفر معلومات التفعيل الأخير (مستخدم قديم)، ننشئ معلومات أساسية من licenseData
    if (!lastInfo) {
        lastInfo = {
            plan: licenseData.plan,
            activatedDate: null, // لا نعرف متى تم التفعيل
            expiry: licenseData.expiry,
            notified: false
        };
    }

    // إذا تم إظهار التنبيه مسبقاً لهذا التفعيل، لا تظهره مرة أخرى
    if (lastInfo.notified) return;

    // إظهار النافذة
    window.showExpirationModal(lastInfo.plan, lastInfo.activatedDate, lastInfo.expiry);

    // تعليم التنبيه كـ "تم الإظهار"
    lastInfo.notified = true;
    localStorage.setItem('last_activation_info', JSON.stringify(lastInfo));
}

window.showExpirationModal = function (plan, activatedDate, expiry) {
    const modal = document.getElementById('expiration-thank-you-modal');
    if (!modal) return;

    const planText = plan === 'PREMIUM' ? 'نسخة كاملة (Premium)' : 'نسخة أساسية (Basic)';

    // حساب المدة
    let durationText = "غير محددة";
    if (activatedDate) {
        const diffMs = expiry - activatedDate;
        const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
        const diffMonths = Math.round(diffDays / 30);

        if (plan === 'PREMIUM') {
            durationText = diffMonths + " شهر" + (diffMonths > 10 ? "" : "اً");
            if (diffMonths === 1) durationText = "شهر واحد";
            if (diffMonths === 2) durationText = "شهرين";
        } else {
            durationText = diffDays + " يوم" + (diffDays > 10 ? "" : "اً");
            if (diffDays === 1) durationText = "يوم واحد";
            if (diffDays === 2) durationText = "يومين";
        }
    } else {
        // إذا لم تتوفر المدة، نظهر نصاً بديلاً
        durationText = plan === 'PREMIUM' ? "عدة أشهر" : "عدة أيام";
    }

    const elPlan = document.getElementById('exp-license-plan');
    const elDuration = document.getElementById('exp-license-duration');

    if (elPlan) elPlan.textContent = planText;
    if (elDuration) elDuration.textContent = durationText;

    modal.classList.add('open');
};

window.closeExpirationModal = function () {
    const modal = document.getElementById('expiration-thank-you-modal');
    if (modal) modal.classList.remove('open');

    // توجيه المستخدم لصفحة التفعيل
    if (typeof window.switchSection === 'function') {
        window.switchSection(null, 'activation');
    }
};

// --- Splash Screen Logic ---
function updateSplashUI(isActivated) {
    const splashLoading = document.getElementById('splash-loading');
    const splashActions = document.getElementById('splash-actions');
    const loginBtn = document.getElementById('splash-login-btn');
    const activateBtn = document.getElementById('splash-activate-btn');

    if (splashLoading) splashLoading.style.display = 'none';
    if (splashActions) splashActions.style.display = 'flex';

    if (isActivated) {
        if (loginBtn) loginBtn.style.display = 'inline-flex';
        if (activateBtn) activateBtn.style.display = 'none';
    } else {
        if (loginBtn) loginBtn.style.display = 'none';
        if (activateBtn) activateBtn.style.display = 'inline-flex';
    }
}

window.enterApp = function () {
    const splash = document.getElementById('splash-screen');
    if (splash) {
        splash.classList.add('fade-out');
        document.body.classList.remove('splash-active');
        setTimeout(() => {
            splash.style.display = 'none';
        }, 500);
    }
};

window.enterActivation = function () {
    enterApp();
    if (typeof window.switchSection === 'function') {
        window.switchSection(null, 'activation');
        // Optional: Open modal directly
        // window.showActivationModal();
    }
};

document.addEventListener('DOMContentLoaded', () => {
    // Delay initActivation slightly to ensure DOM is ready and show splash for at least a moment
    setTimeout(() => {
        initActivation();
    }, 1500); // 1.5 seconds minimum splash time

    // Splash Logo Interaction
    const splashLogo = document.querySelector('.splash-logo img');
    if (splashLogo) {
        splashLogo.addEventListener('mouseenter', () => {
            playSplashSound();
        });
    }
});

// --- Sound Logic ---
function playSplashSound() {
    try {
        const AudioContext = window.AudioContext || window.webkitAudioContext;
        if (!AudioContext) return;

        const ctx = new AudioContext();
        const oscillator = ctx.createOscillator();
        const gainNode = ctx.createGain();

        oscillator.type = 'sine';
        // A nice "magical" chime chord arpeggio effect
        const now = ctx.currentTime;

        // Base tone
        oscillator.frequency.setValueAtTime(523.25, now); // C5
        oscillator.frequency.exponentialRampToValueAtTime(1046.50, now + 0.1); // C6 slide up

        gainNode.gain.setValueAtTime(0.1, now);
        gainNode.gain.exponentialRampToValueAtTime(0.01, now + 0.5);

        oscillator.connect(gainNode);
        gainNode.connect(ctx.destination);

        oscillator.start();
        oscillator.stop(now + 0.5);

    } catch (e) {
        console.warn('Audio play failed', e);
    }
}

