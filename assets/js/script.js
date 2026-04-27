// ============================================
// NOVA LIVING — Store Interactions
// ============================================

(function () {
    'use strict';

    // ===== Bundle Selection & Live Summary =====
    const bundleCards = document.querySelectorAll('.bundle-card');
    const sumProduct = document.getElementById('sumProduct');
    const sumTotal = document.getElementById('sumTotal');

    const productLabels = {
        single: 'قطعة واحدة NOVA',
        duo: 'قطعتين NOVA',
        trio: '3 قطع NOVA + حامل ذهبي',
    };

    function updateSummary() {
        const checked = document.querySelector('input[name="bundle"]:checked');
        if (!checked) return;
        const card = checked.closest('.bundle-card');
        const price = card.dataset.price;
        const value = checked.value;

        if (sumProduct) sumProduct.textContent = productLabels[value] || '—';
        if (sumTotal) sumTotal.textContent = `${price} د.إ`;
    }

    bundleCards.forEach((card) => {
        card.addEventListener('click', () => {
            const input = card.querySelector('input');
            if (input) {
                input.checked = true;
                updateSummary();
            }
        });
    });

    updateSummary();

    // ===== Gallery Switcher =====
    const galleryMain = document.getElementById('galleryMain');
    const thumbs = document.querySelectorAll('.thumb');

    const galleryLabels = {
        warm: '🌅 وضع الغروب',
        rose: '🌸 الوضع الوردي',
        purple: '💜 الوضع الأرجواني',
        gold: '✨ الوضع الذهبي',
        room: '🛏️ في الغرفة',
        box: '📦 التغليف الفاخر',
    };

    thumbs.forEach((thumb) => {
        thumb.addEventListener('click', () => {
            const target = thumb.dataset.img;
            thumbs.forEach((t) => t.classList.remove('active'));
            thumb.classList.add('active');

            if (galleryMain) {
                const img = galleryMain.querySelector('.gallery-img');
                const label = galleryMain.querySelector('.gallery-label');
                if (img) {
                    img.style.opacity = '0';
                    setTimeout(() => {
                        img.className = `gallery-img mood-${target}`;
                        if (target === 'box') {
                            img.innerHTML = '<span class="gallery-label">' + galleryLabels[target] + '</span>';
                        } else {
                            img.innerHTML = `
                                <div class="gallery-lamp">
                                    <div class="gl-glow"></div>
                                    <div class="gl-body"></div>
                                    <div class="gl-base"></div>
                                </div>
                                <span class="gallery-label">${galleryLabels[target]}</span>`;
                        }
                        img.style.opacity = '1';
                    }, 200);
                }
            }
        });
    });

    // ===== Video Demo Click =====
    const mainVideo = document.getElementById('mainVideo');
    if (mainVideo) {
        mainVideo.addEventListener('click', () => {
            // Placeholder — will be replaced with real video later
            alert('🎬 الفيديو الحقيقي راح يضاف هنا قريباً!\n\nلتجربة المنتج، اطلب الآن - الدفع عند الاستلام.');
            const orderSection = document.getElementById('order');
            if (orderSection) {
                const top = orderSection.getBoundingClientRect().top + window.pageYOffset - 80;
                window.scrollTo({ top, behavior: 'smooth' });
            }
        });
    }

    // ===== TikTok Cards Click =====
    document.querySelectorAll('.tt-card').forEach((card) => {
        card.addEventListener('click', () => {
            window.open('https://www.tiktok.com/search?q=sunset+lamp', '_blank');
        });
    });

    // ===== Form Submission =====
    const form = document.getElementById('orderForm');
    const modal = document.getElementById('successModal');

    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();

            // Simple validation
            if (!form.checkValidity()) {
                form.reportValidity();
                return;
            }

            // Collect order data (would POST to backend in production)
            const data = {
                name: form.name.value,
                phone: form.phone.value,
                emirate: form.emirate.value,
                city: form.city.value,
                address: form.address.value,
                notes: form.notes.value,
                bundle: document.querySelector('input[name="bundle"]:checked')?.value,
                total: sumTotal?.textContent,
                timestamp: new Date().toISOString(),
            };

            console.log('Order submitted:', data);

            // Show success modal
            if (modal) {
                modal.classList.add('active');
                document.body.style.overflow = 'hidden';
            }

            form.reset();
            updateSummary();
        });
    }

    // Phone input — auto-format and limit to digits
    const phoneInput = document.getElementById('phone');
    if (phoneInput) {
        phoneInput.addEventListener('input', (e) => {
            e.target.value = e.target.value.replace(/[^\d+]/g, '');
        });
    }

    // ===== Modal Close =====
    window.closeModal = function () {
        if (modal) {
            modal.classList.remove('active');
            document.body.style.overflow = '';
        }
    };

    if (modal) {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal();
        });
    }

    // ===== Smooth Scroll =====
    document.querySelectorAll('a[href^="#"]').forEach((link) => {
        link.addEventListener('click', (e) => {
            const target = link.getAttribute('href');
            if (target === '#' || target.length < 2) return;
            const el = document.querySelector(target);
            if (!el) return;
            e.preventDefault();
            const headerOffset = 80;
            const top = el.getBoundingClientRect().top + window.pageYOffset - headerOffset;
            window.scrollTo({ top, behavior: 'smooth' });
        });
    });

    // ===== Reveal On Scroll =====
    const revealElements = document.querySelectorAll('.feature-card, .review-card, .bundle-card, .ba-card');

    if ('IntersectionObserver' in window) {
        const observer = new IntersectionObserver(
            (entries) => {
                entries.forEach((entry) => {
                    if (entry.isIntersecting) {
                        entry.target.style.opacity = '1';
                        entry.target.style.transform = 'translateY(0)';
                        observer.unobserve(entry.target);
                    }
                });
            },
            { threshold: 0.1, rootMargin: '0px 0px -50px 0px' }
        );

        revealElements.forEach((el) => {
            el.style.opacity = '0';
            el.style.transform = 'translateY(20px)';
            el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
            observer.observe(el);
        });
    }

    // ===== Live Stock Counter (FOMO) =====
    const stockBanner = document.createElement('div');
    const remaining = Math.floor(Math.random() * 12) + 7;

    // ===== Recent Activity Toast (Social Proof) =====
    const recentBuyers = [
        { name: 'فاطمة', city: 'دبي', time: 'منذ 3 دقائق' },
        { name: 'أحمد', city: 'أبوظبي', time: 'منذ 7 دقائق' },
        { name: 'مريم', city: 'الشارقة', time: 'منذ 12 دقيقة' },
        { name: 'خالد', city: 'دبي', time: 'منذ 18 دقيقة' },
        { name: 'نورة', city: 'العين', time: 'منذ 24 دقيقة' },
    ];

    const toast = document.createElement('div');
    toast.className = 'activity-toast';
    toast.style.cssText = `
        position: fixed;
        bottom: 90px;
        right: 24px;
        background: rgba(26, 26, 26, 0.95);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 14px;
        padding: 12px 16px;
        display: flex;
        align-items: center;
        gap: 12px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
        z-index: 80;
        max-width: 280px;
        font-size: 13px;
        color: #f5f5f5;
        opacity: 0;
        transform: translateY(20px);
        transition: all 0.4s ease;
        pointer-events: none;
    `;

    document.body.appendChild(toast);

    function showActivity() {
        const buyer = recentBuyers[Math.floor(Math.random() * recentBuyers.length)];
        toast.innerHTML = `
            <div style="font-size:24px">🎉</div>
            <div style="line-height:1.4">
                <strong style="color:#d4af37;display:block;font-size:13px">${buyer.name} من ${buyer.city}</strong>
                <small style="color:#888;font-size:11px">طلب مصباح NOVA · ${buyer.time}</small>
            </div>
        `;
        toast.style.opacity = '1';
        toast.style.transform = 'translateY(0)';

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateY(20px)';
        }, 4500);
    }

    // First toast after 6s, then every 18-25s
    setTimeout(() => {
        showActivity();
        setInterval(showActivity, 22000);
    }, 6000);
})();
