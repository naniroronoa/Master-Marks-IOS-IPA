const LICENSE_SECRET = "TEACHER_XP_2026_SECRET";

/**
 * Generates a key: TLILI - (8 chars) + 50 random-looking chars.
 */
function generateGHKey(hwid, plan, expiry) {
    const prefix = "TLILI - ";
    const bodyLength = 50;

    // 1. Salt (8 chars)
    const saltChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let salt = "";
    for (let i = 0; i < 8; i++) salt += saltChars[Math.floor(Math.random() * saltChars.length)];

    // 2. Data pack
    const hwidDigits = hwid.replace(/\D/g, '');
    const planChar = plan === 'PREMIUM' ? 'P' : 'B';
    const data = `${salt}|${hwidDigits}|${planChar}|${expiry}`;

    // 3. XOR with fixed secret (Salt is already inside the data)
    const encrypted = xorEncrypt(data, LICENSE_SECRET);

    // 4. Encode to 50 chars using a wide alphabet
    // We'll use Base64 as an intermediate, then map its chars to a wider set
    const b64 = btoa(encrypted).replace(/=/g, '');

    const targetAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}|;:,.<>?";

    // To make it look "random" and use symbols equally:
    // We map the B64 string index-by-index + add random noise at the end
    let body = "";
    for (let i = 0; i < b64.length; i++) {
        // Simple 1-to-1 mapping from B64 set to sub-range of targetAlphabet
        // B64 is 64 chars. targetAlphabet is ~80+.
        const b64Index = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/".indexOf(b64[i]);
        body += targetAlphabet[b64Index];
    }

    // Pad with purely random symbols from the remaining part of the targetAlphabet
    while (body.length < bodyLength) {
        body += targetAlphabet[Math.floor(Math.random() * targetAlphabet.length)];
    }

    return prefix + body;
}

// Keeping these for potential metadata storage or legacy compatibility if needed
function xorEncrypt(str, key) {
    let result = '';
    for (let i = 0; i < str.length; i++) {
        result += String.fromCharCode(str.charCodeAt(i) ^ key.charCodeAt(i % key.length));
    }
    return result;
}

function xorDecrypt(str, key) {
    return xorEncrypt(str, key);
}
