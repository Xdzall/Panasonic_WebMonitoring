/**
 * Layout.js - MonitoringSystem
 * Menggabungkan semua functionality dari _Layout.cshtml
 */

(function ($) {
    'use strict';

    /**
     * 1. ACTIVE NAVIGATION HIGHLIGHT
     * Menandai menu yang aktif berdasarkan URL saat ini
     */
    function initActiveNavigation() {
        // Ambil path saat ini dan bersihkan
        var currentPath = window.location.pathname.toLowerCase();

        // Normalisasi currentPath: hilangkan '/index' di akhir jika ada
        if (currentPath.endsWith('/index')) {
            currentPath = currentPath.substring(0, currentPath.length - 6);
        }
        // Jika kosong setelah dihapus index (misal root), jadikan '/'
        if (currentPath === '') {
            currentPath = '/';
        }

        // Reset semua active class
        $('#bdSidebar .nav-item a').removeClass('active');

        // Cek setiap link
        $('#bdSidebar .nav-item a').each(function () {
            var $this = $(this);
            var href = $this.attr('href');

            // Skip jika href tidak valid
            if (!href || href === '#') return;

            // Bersihkan href dari '~' dan '/index'
            var linkPath = href.replace('~', '').toLowerCase();

            if (linkPath.endsWith('/index')) {
                linkPath = linkPath.substring(0, linkPath.length - 6);
            }
            if (linkPath === '') {
                linkPath = '/';
            }

            // EXACT MATCH atau SUB-PAGE
            if (currentPath === linkPath) {
                $this.addClass('active');
            }
            // Cek apakah ini sub-halaman
            // Tambah '/' setelah linkPath agar tidak salah baca folder yang mirip
            else if (currentPath.startsWith(linkPath + '/')) {
                $this.addClass('active');
            }
        });
    }

    /**
     * 2. DATE & TIME UPDATE
     * Update tanggal dan waktu real-time
     */
    function updateDateTime() {
        const now = new Date();
        const dateStr = now.toLocaleDateString('id-ID', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        }).replace(/\//g, '.');

        const timeStr = now.toLocaleTimeString('id-ID', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: false
        });

        document.getElementById('current-date').textContent = dateStr;
        document.getElementById('current-time').textContent = timeStr;
    }

    function initDateTime() {
        updateDateTime();
        setInterval(updateDateTime, 1000);
    }

    /**
     * 3. MORE MENU TOGGLE
     * Toggle popup menu untuk "More" button
     */
    function initMoreMenu() {
        const moreBtn = document.getElementById('moreMenuBtn');
        const morePopup = document.getElementById('moreMenuPopup');

        if (moreBtn && morePopup) {
            // Toggle menu saat button diklik
            moreBtn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();

                // Toggle display
                if (morePopup.style.display === 'none' || morePopup.style.display === '') {
                    morePopup.style.display = 'block';
                } else {
                    morePopup.style.display = 'none';
                }
            });

            // Close popup ketika klik di luar
            document.addEventListener('click', function (e) {
                if (!moreBtn.contains(e.target) && !morePopup.contains(e.target)) {
                    morePopup.style.display = 'none';
                }
            });
        }
    }

    /**
     * INITIALIZATION
     * Jalankan semua fungsi saat DOM ready
     */
    $(document).ready(function () {
        initActiveNavigation();
        initDateTime();
    });

    // More menu harus diinit setelah DOM fully loaded
    document.addEventListener('DOMContentLoaded', function () {
        initMoreMenu();
    });

})(jQuery);
