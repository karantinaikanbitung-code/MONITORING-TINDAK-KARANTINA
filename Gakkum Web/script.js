document.addEventListener('DOMContentLoaded', () => {
    console.log("Web Monitoring Dashboard - v1.2 (Dynamic Title + Vol XLS)");
    // Real Data from Excel files
    const tableData = [
        // Penahanan (2).xls
        { id: 1, noDok: '2026-H2.0-7101.0-K.1.1-000007', tgl: '2026-01-22', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor', volP0: '16,000', volP5: '16,000', satuan: 'kemasan' },
        { id: 2, noDok: '2026-H2.0-7101.0-K.1.1-000003', tgl: '2026-01-07', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Ayam', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor', volP0: '24,000', volP5: '24,000', satuan: 'kemasan' },
        { id: 3, noDok: '2026-H2.0-7101.0-K.1.1-000002', tgl: '2026-01-05', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor', volP0: '16,000', volP5: '16,000', satuan: 'kemasan' },
        { id: 4, noDok: '2026-H2.0-7101.0-K.1.1-000001', tgl: '2026-01-01', pengirim: 'NN', penerima: 'NN', asal: 'PHILIPINA', tujuan: 'INDONESIA - KOTA BITUNG', komoditas: 'Unggas Besar Kesayangan', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor', volP0: '218,000', volP5: '218,000', satuan: 'ekor' },
        { id: 5, noDok: '2026-H2.0-7101.0-K.1.1-000005', tgl: '2026-01-14', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Sapi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor', volP0: '6,000', volP5: '6,000', satuan: 'kemasan' },
        { id: 6, noDok: '2026-H2.0-7101.0-K.1.1-000006', tgl: '2026-01-15', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor', volP0: '8,000', volP5: '8,000', satuan: 'kemasan' },
        { id: 7, noDok: '2026-H2.0-7101.0-K.1.1-000004', tgl: '2026-01-13', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor', volP0: '90,000', volP5: '90,000', satuan: 'kemasan' },

        // Penahanan (3).xls
        { id: 8, noDok: '2026-I2.0-7101.0-K.1.1-000004', tgl: '2026-01-30', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Ikan Teri', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor', volP0: '0,500', volP5: '0,500', satuan: 'kilogram' },
        { id: 9, noDok: '2026-I2.0-7101.0-K.1.1-000002', tgl: '2026-01-20', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Ikan Patin', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor', volP0: '5,500', volP5: '5,500', satuan: 'kilogram' },
        { id: 10, noDok: '2026-I2.0-7101.0-K.1.1-000001', tgl: '2026-01-09', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Sidat', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor', volP0: '0,300', volP5: '0,300', satuan: 'kilogram' },
        { id: 11, noDok: '2026-I2.0-7101.0-K.1.1-000003', tgl: '2026-01-28', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Ikan Teri', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor', volP0: '9,200', volP5: '9,200', satuan: 'kilogram' },

        // Penahanan (4).xls
        { id: 12, noDok: '2026-T2.0-7101.0-K.1.1-000004', tgl: '2026-01-20', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'CABE KERING', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor', volP0: '1,000', volP5: '0,250', satuan: 'kemasan' },
        { id: 13, noDok: '2026-T2.0-7101.0-K.1.1-000002', tgl: '2026-01-15', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'JAMUR KERING', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor', volP0: '2,000', volP5: '2,000', satuan: 'kemasan' },
        { id: 14, noDok: '2026-T2.0-7101.0-K.1.1-000001', tgl: '2026-01-06', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BUAH-BUAHAN', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor', volP0: '1,000', volP5: '1,000', satuan: 'kemasan' },
        { id: 15, noDok: '2026-T2.0-7101.0-K.1.1-000007', tgl: '2026-01-31', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BAWANG DAUN', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor', volP0: '0,600', volP5: '0,600', satuan: 'kilogram' },
        { id: 16, noDok: '2026-T2.0-7101.0-K.1.1-000006', tgl: '2026-01-22', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BUAH-BUAHAN', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor', volP0: '1,000', volP5: '1,000', satuan: 'kemasan' },
        { id: 17, noDok: '2026-T2.0-7101.0-K.1.1-000005', tgl: '2026-01-21', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BIDARA/JUJUBE', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor', volP0: '2,000', volP5: '2,000', satuan: 'kemasan' },
        { id: 18, noDok: '2026-T2.0-7101.1-K.1.1-000001', tgl: '2026-01-05', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BIJI CHIA', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor', volP0: '1,000', volP5: '1,000', satuan: 'kemasan' },

        // Penahanan (5).xls
        { id: 19, noDok: '2026-H4.0-7103.0-K.1.1-000039', tgl: '2026-01-11', pengirim: 'PINDI KURNIAWAN', penerima: 'AGUNG WAHYU PRIBADI', asal: 'KOTA MANADO', tujuan: 'KEPULAUAN SANGIHE', komoditas: 'Burung Parkit', tkp: 'BKHIT Sulawesi Utara - Pelabuhan Laut Tahuna', ket: 'Lengkap', category: 'Hewan', permohonan: 'Domestik Masuk', volP0: '10,000', volP5: '10,000', satuan: 'ekor' },
        { id: 20, noDok: '2026-H4.0-7101.1-K.1.1-000008', tgl: '2026-01-08', pengirim: 'MERLIN MATAHARI', penerima: 'MERLIN MATAHARI', asal: 'KOTA TERNATE', tujuan: 'KOTA MANADO', komoditas: 'Kucing', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Domestik Masuk', volP0: '3,000', volP5: '3,000', satuan: 'ekor' },
        { id: 21, noDok: '2026-H4.0-7101.1-K.1.1-000005', tgl: '2026-01-06', pengirim: 'Aldaier Makatindu', penerima: 'Jouce Kansil', asal: 'KEPULAUAN SANGIHE', tujuan: 'KOTA MANADO', komoditas: 'Ayam', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Domestik Masuk', volP0: '3,000', volP5: '3,000', satuan: 'ekor' },
    ];

    const tableBody = document.getElementById('table-body');
    const modal = document.getElementById('detail-modal');
    const closeModalBtn = document.querySelector('.close-modal');
    const previewContainer = document.getElementById('preview-container');
    const fileViewer = document.getElementById('file-viewer');
    const pdfFrame = document.getElementById('pdf-frame');
    const viewerFilename = document.getElementById('viewer-filename');
    const WORKER_URL = 'https://gakkum-monitoring-api.karantinaikanbitung.workers.dev';
    const boxPemusnahan = document.getElementById('box-pemusnahan');
    let currentUploadContext = null;
    let currentDocId = null;
    let uploadedFiles = {};

    // Render Table
    function renderTable(data) {
        tableBody.innerHTML = '';
        data.forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${row.noDok}</td>
                <td>${row.tgl}</td>
                <td>${row.category}</td>
                <td>${row.permohonan}</td>
                <td>${row.pengirim}</td>
                <td>${row.penerima}</td>
                <td>${row.asal}</td>
                <td>${row.tujuan}</td>
                <td>${row.komoditas}</td>
                <td>${row.tkp}</td>
                <td>${row.ket}</td>
                <td>
                    <button class="btn btn-primary btn-sm" onclick="openDetail(${row.id})">
                        <i class="fa-solid fa-eye"></i> Detail
                    </button>
                </td>
            `;
            tableBody.appendChild(tr);
        });
    }

    // Modal Logic
    window.openDetail = (id) => {
        currentDocId = id;
        localStorage.setItem('lastDocId', id);

        // Update Modal Title with Doc Suffix
        const modalTitle = document.getElementById('modal-title');
        const doc = tableData.find(d => d.id === id);
        if (modalTitle && doc) {
            // Get suffix: everything after the last 3 hyphens (e.g. K.1.1-000007)
            // Format: YYYY-CATEGORY-CODE-SUFFIX-NUMBER -> we want SUFFIX-NUMBER
            const parts = doc.noDok.split('-');
            const suffix = parts.length > 2 ? parts.slice(-2).join('-') : doc.noDok;
            modalTitle.textContent = `Detail Dokumen ${suffix}`;
        }

        modal.style.display = 'flex';
        resetView();
        checkExistingFiles(id);
    };

    async function checkExistingFiles(id) {
        // First, try to load from localStorage for immediate UI feedback
        const savedData = localStorage.getItem(`files_${id}`);
        if (savedData) {
            uploadedFiles = JSON.parse(savedData);
            Object.keys(uploadedFiles).forEach(ctx => {
                updateUIIndicator(ctx, true);
            });
        }

        const contexts = ['LI', 'Kronologi', 'PUL BAKET', 'Penolakan', 'Pemusnahan'];
        for (const ctx of contexts) {
            try {
                const response = await fetch(`${WORKER_URL}/check/${id}/${ctx}`);
                if (!response.ok) {
                    // If not found on server, but we have a local flag or it was deleted, sync
                    continue;
                }

                const data = await response.json();
                if (data.exists) {
                    uploadedFiles[ctx] = {
                        name: data.filename,
                        url: `${WORKER_URL}/view/${id}/${ctx}/${data.filename}`
                    };
                    updateUIIndicator(ctx, true);
                } else if (uploadedFiles[ctx] && !uploadedFiles[ctx].isLocal) {
                    // Sync: if server says it doesn't exist and we thought it was a remote file, remove it
                    delete uploadedFiles[ctx];
                    updateUIIndicator(ctx, false);
                }
            } catch (err) {
                console.error(`Error checking ${ctx}:`, err);
            }
        }
        // Save back to localStorage after remote check
        localStorage.setItem(`files_${id}`, JSON.stringify(uploadedFiles));

        // After checking files, if there's a saved context, load it
        const savedCtx = localStorage.getItem('lastContext');
        if (savedCtx && id === parseInt(localStorage.getItem('lastDocId'))) {
            loadPreview(savedCtx);
            // Open submenu if needed
            if (['LI', 'Kronologi', 'PUL BAKET'].includes(savedCtx)) {
                const submenu = document.getElementById('submenu-penahanan');
                if (submenu) submenu.style.display = 'block';
            }
        }
    }

    function updateUIIndicator(ctx, exists) {
        const boxMap = {
            'Pemusnahan': 'box-pemusnahan',
            'Penolakan': 'box-penolakan',
            'LI': 'item-LI',
            'Kronologi': 'item-Kronologi',
            'PUL BAKET': 'item-PUL-BAKET'
        };

        const elementId = boxMap[ctx];
        if (!elementId) return;

        const el = document.getElementById(elementId);
        if (!el) return;

        if (exists) {
            el.classList.add('file-available');
            if (el.classList.contains('action-box')) {
                el.classList.remove('red-box');
                el.classList.add('blue-box');
            } else if (el.classList.contains('submenu-item')) {
                el.classList.add('has-file');
            }
        } else {
            el.classList.remove('file-available');
            if (el.classList.contains('action-box')) {
                if (ctx === 'Pemusnahan') {
                    el.classList.remove('blue-box');
                    el.classList.add('red-box');
                }
            } else if (el.classList.contains('submenu-item')) {
                el.classList.remove('has-file');
            }
        }

        // If this is the active context, refresh preview if it exists
        if (ctx === currentUploadContext) {
            if (exists && uploadedFiles[ctx]) {
                showPdf(uploadedFiles[ctx].name, uploadedFiles[ctx].url);
            } else {
                previewContainer.classList.remove('hidden');
                fileViewer.classList.add('hidden');
            }
        }
    }

    closeModalBtn.addEventListener('click', () => {
        modal.style.display = 'none';
        localStorage.removeItem('lastDocId');
        localStorage.removeItem('lastContext');
        // Reset Title
        const modalTitle = document.getElementById('modal-title');
        if (modalTitle) modalTitle.textContent = 'Detail Dokumen';
    });

    window.onclick = (event) => {
        if (event.target === modal) {
            modal.style.display = 'none';
            localStorage.removeItem('lastDocId');
            localStorage.removeItem('lastContext');
            // Reset Title
            const modalTitle = document.getElementById('modal-title');
            if (modalTitle) modalTitle.textContent = 'Detail Dokumen';
        }
    };

    function resetView() {
        previewContainer.classList.remove('hidden');
        fileViewer.classList.add('hidden');
        uploadedFiles = {};
        // Reset submenus and boxes
        document.querySelectorAll('.submenu-item').forEach(el => el.classList.remove('has-file'));
        document.querySelectorAll('.file-available').forEach(el => el.classList.remove('file-available'));
        if (boxPemusnahan) {
            boxPemusnahan.classList.remove('blue-box');
            boxPemusnahan.classList.add('red-box');
        }
    }

    // Sidebar Interactions
    window.toggleSubmenu = (id) => {
        const submenu = document.getElementById(id);
        submenu.style.display = submenu.style.display === 'block' ? 'none' : 'block';
    };

    window.loadPreview = (type) => {
        currentUploadContext = type;
        localStorage.setItem('lastContext', type);
        if (uploadedFiles[type]) {
            showPdf(uploadedFiles[type].name, uploadedFiles[type].url);
        } else {
            previewContainer.classList.remove('hidden');
            fileViewer.classList.add('hidden');
            const emptyStateP = document.querySelector('.empty-state p');
            if (emptyStateP) emptyStateP.textContent = `Upload file untuk ${type}`;
        }
    };

    window.triggerFileUpload = () => {
        document.getElementById('file-input').click();
    };

    document.getElementById('file-input').addEventListener('change', async (e) => {
        if (e.target.files.length > 0) {
            const file = e.target.files[0];

            // 1. Immediate Local Preview
            const localURL = URL.createObjectURL(file);
            showPdf(file.name, localURL);

            // Update UI indicator immediately
            updateUIIndicator(currentUploadContext, true);

            // 2. Attempt Background Upload
            const emptyStateText = document.querySelector('.empty-state p');
            const originalText = emptyStateText ? emptyStateText.textContent : "";
            if (emptyStateText) emptyStateText.textContent = "Uploading to server...";

            try {
                const uploadUrl = `${WORKER_URL}/upload/${currentDocId}/${currentUploadContext}/${encodeURIComponent(file.name)}`;
                const response = await fetch(uploadUrl, {
                    method: 'PUT',
                    body: file,
                    headers: { 'Content-Type': file.type }
                });

                if (response.ok) {
                    const remoteURL = `${WORKER_URL}/view/${currentDocId}/${currentUploadContext}/${encodeURIComponent(file.name)}`;
                    uploadedFiles[currentUploadContext] = {
                        name: file.name,
                        url: remoteURL
                    };
                    // Persist to localStorage
                    localStorage.setItem(`files_${currentDocId}`, JSON.stringify(uploadedFiles));

                    // Update iframe to remote URL (cleaner)
                    pdfFrame.src = remoteURL;
                } else {
                    console.warn("Upload failed but file is shown locally:", response.statusText);
                    // Still keep the local preview for the session
                    uploadedFiles[currentUploadContext] = {
                        name: file.name,
                        url: localURL,
                        isLocal: true
                    };
                    localStorage.setItem(`files_${currentDocId}`, JSON.stringify(uploadedFiles));
                    if (emptyStateText) emptyStateText.textContent = originalText;
                }
            } catch (err) {
                console.error("Upload error:", err);
                // Notification instead of blocking alert
                const notify = document.createElement('div');
                notify.style = "position:fixed; bottom:20px; right:20px; background:#f44336; color:white; padding:10px; border-radius:5px; z-index:2000;";
                notify.textContent = "Gagal upload ke server, tapi file tampil lokal.";
                document.body.appendChild(notify);
                setTimeout(() => notify.remove(), 3000);

                uploadedFiles[currentUploadContext] = {
                    name: file.name,
                    url: localURL,
                    isLocal: true
                };
                localStorage.setItem(`files_${currentDocId}`, JSON.stringify(uploadedFiles));
                if (emptyStateText) emptyStateText.textContent = originalText;
            }
        }
    });

    function showPdf(filename, fileUrl) {
        previewContainer.classList.add('hidden');
        fileViewer.classList.remove('hidden');
        viewerFilename.textContent = filename;
        pdfFrame.src = fileUrl;
    }

    window.resetUpload = async () => {
        if (!currentUploadContext || !uploadedFiles[currentUploadContext]) return;

        const ctx = currentUploadContext;
        const fileToDelete = uploadedFiles[ctx];

        const confirmDelete = confirm(`Hapus file ${fileToDelete.name} dari server?`);
        if (!confirmDelete) return;

        try {
            const deleteUrl = `${WORKER_URL}/delete/${currentDocId}/${ctx}/${encodeURIComponent(fileToDelete.name)}`;
            const response = await fetch(deleteUrl, { method: 'DELETE' });

            if (response.ok) {
                delete uploadedFiles[ctx];
                localStorage.setItem(`files_${currentDocId}`, JSON.stringify(uploadedFiles));
                updateUIIndicator(ctx, false);

                // Show empty state
                previewContainer.classList.remove('hidden');
                fileViewer.classList.add('hidden');
                document.getElementById('file-input').value = '';
                const emptyStateP = document.querySelector('.empty-state p');
                if (emptyStateP) emptyStateP.textContent = `Upload file untuk ${ctx}`;
            } else {
                alert("Gagal menghapus file dari server: " + response.statusText);
            }
        } catch (err) {
            console.error("Delete error:", err);
            alert("Terjadi kesalahan saat menghapus file.");
        }
    };

    // Theme Toggle
    const themeToggleBtn = document.getElementById('theme-toggle');
    themeToggleBtn.addEventListener('click', () => {
        document.body.classList.toggle('dark-mode');
        const icon = themeToggleBtn.querySelector('i');
        if (document.body.classList.contains('dark-mode')) {
            icon.classList.remove('fa-moon');
            icon.classList.add('fa-sun');
        } else {
            icon.classList.remove('fa-sun');
            icon.classList.add('fa-moon');
        }
    });

    // Profile Logout
    document.getElementById('profile-btn').addEventListener('click', () => {
        const dropdown = document.getElementById('profile-dropdown');
        dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
    });

    document.getElementById('logout-btn').addEventListener('click', (e) => {
        e.preventDefault();
        alert('Logging out...');
        // window.location.href = 'login.html';
    });

    // Filter Logic
    window.filterTable = () => {
        const karantina = document.getElementById('filter-karantina').value;
        const permohonan = document.getElementById('filter-permohonan').value;
        const tglAwal = document.getElementById('filter-tgl-awal').value;
        const tglAkhir = document.getElementById('filter-tgl-akhir').value;

        const filteredData = tableData.filter(row => {
            const matchKarantina = karantina === '' || row.category === karantina;
            const matchPermohonan = permohonan === '' || row.permohonan === permohonan;

            let matchDate = true;
            if (tglAwal && tglAkhir) {
                matchDate = row.tgl >= tglAwal && row.tgl <= tglAkhir;
            } else if (tglAwal) {
                matchDate = row.tgl >= tglAwal;
            } else if (tglAkhir) {
                matchDate = row.tgl <= tglAkhir;
            }

            return matchKarantina && matchPermohonan && matchDate;
        });

        renderTable(filteredData);
    };

    window.exportXLS = () => {
        // Re-filter to get current view data
        const karantinaVal = document.getElementById('filter-karantina').value;
        const permohonanVal = document.getElementById('filter-permohonan').value;
        const tglAwal = document.getElementById('filter-tgl-awal').value;
        const tglAkhir = document.getElementById('filter-tgl-akhir').value;

        const filteredData = tableData.filter(row => {
            const matchKarantina = karantinaVal === '' || row.category === karantinaVal;
            const matchPermohonan = permohonanVal === '' || row.permohonan === permohonanVal;
            let matchDate = true;
            if (tglAwal && tglAkhir) { matchDate = row.tgl >= tglAwal && row.tgl <= tglAkhir; }
            else if (tglAwal) { matchDate = row.tgl >= tglAwal; }
            else if (tglAkhir) { matchDate = row.tgl <= tglAkhir; }
            return matchKarantina && matchPermohonan && matchDate;
        });

        // Header Information Helpers
        const karantinaMap = { 'Hewan': 'KH', 'Ikan': 'KI', 'Tumbuhan': 'KT' };
        const permohonanMap = {
            'Ekspor': 'EK',
            'Impor': 'IM',
            'Domestik Keluar': 'DK',
            'Domestik Masuk': 'DM'
        };

        const karantinaText = karantinaMap[karantinaVal] || 'SEMUA';
        const permohonanText = permohonanMap[permohonanVal] || 'SEMUA';
        const tglText = (tglAwal && tglAkhir) ? `${tglAwal} S/D ${tglAkhir}` : (tglAwal ? `MULAI ${tglAwal}` : (tglAkhir ? `HINGGA ${tglAkhir}` : 'SEMUA PERIODE'));
        const now = new Date();
        const waktuCetak = now.toISOString().slice(0, 19).replace('T', ' ');

        // Generate HTML for Excel with Header and 18 Columns
        let tableHTML = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body { font-family: Arial, sans-serif; }
                    table { border-collapse: collapse; width: 100%; margin-top: 20px; }
                    th, td { border: 1px solid #000; padding: 5px; text-align: left; }
                    center { text-align: center; }
                    h4 { margin: 5px 0; }
                </style>
            </head>
            <body>
                <center>
                    <h4>LAPORAN PENAHANAN ${karantinaText}</h4>
                    <h4>PERMOHONAN ${permohonanText}</h4>
                    <h4>7100</h4>
                    <h4>PERIODE ${tglText}</h4>
                    <h4>Waktu Cetak ${waktuCetak} &nbsp; Oleh DASHBOARD MONITORING</h4>
                    <br><br>
                </center>
                <table>
                    <thead>
                        <tr>
                            <th>No. Dokumen</th>
                            <th>Tgl Dokumen</th>
                            <th>No. P5</th>
                            <th>Tgl P5</th>
                            <th>Satpel</th>
                            <th>Pengirim</th>
                            <th>Penerima</th>
                            <th>Asal</th>
                            <th>Tujuan</th>
                            <th>Alasan</th>
                            <th>Rekomendasi</th>
                            <th>Petugas</th>
                            <th>Komoditas</th>
                            <th>Nama Tercetak</th>
                            <th>Kode HS</th>
                            <th>volume P0</th>
                            <th>volume P5</th>
                            <th>satuan</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        filteredData.forEach(row => {
            tableHTML += `
                <tr>
                    <td>${row.noDok}</td>
                    <td>${row.tgl}</td>
                    <td>-</td>
                    <td>-</td>
                    <td>${row.tkp}</td>
                    <td>${row.pengirim}</td>
                    <td>${row.penerima}</td>
                    <td>${row.asal}</td>
                    <td>${row.tujuan}</td>
                    <td>Media pembawa tidak dilaporkan kepada pejabat karantina...</td>
                    <td>dilakukan penolakan</td>
                    <td>SISTEM MONITORING</td>
                    <td>${row.komoditas}</td>
                    <td>${row.komoditas}</td>
                    <td>00000000</td>
                    <td>${row.volP0}</td>
                    <td>${row.volP5}</td>
                    <td>${row.satuan}</td>
                </tr>
            `;
        });

        tableHTML += `
                    </tbody>
                </table>
            </body>
            </html>
        `;

        // Create Blob and Download
        const blob = new Blob([tableHTML], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Laporan_Penahanan_${karantinaText}_${permohonanText}.xls`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    };

    // Initial Render
    filterTable();

    // Restore Session
    const lastDocId = localStorage.getItem('lastDocId');
    if (lastDocId) {
        openDetail(parseInt(lastDocId));
    }
});
