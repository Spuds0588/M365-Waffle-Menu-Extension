// A self-invoking function to avoid polluting the global namespace
(function() {
    // We need a robust way to wait for the navigation header to exist,
    // as it's rendered by React after the page loads.
    const findNavHeader = setInterval(() => {
        const navHeader = document.querySelector('header.fui-NavDrawerHeader');
        if (navHeader) {
            clearInterval(findNavHeader); // Stop checking once we've found it
            addWaffleButton(navHeader);
        }
    }, 500); // Check every 500ms

    function addWaffleButton(navHeader) {
        // Prevent adding the button if it's already there
        if (document.getElementById('waffle-menu-button')) return;

        const waffleButton = document.createElement('button');
        waffleButton.id = 'waffle-menu-button';
        waffleButton.type = 'button';
        waffleButton.setAttribute('aria-label', 'Open M365 Apps');
        
        // Use theme detection to set the correct icon
        const isDarkMode = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
        const iconUrl = chrome.runtime.getURL(isDarkMode ? "icons/waffle-white.svg" : "icons/waffle-black.svg");
        waffleButton.style.cssText = `
            background: url(${iconUrl}) no-repeat center center;
            background-size: 24px;
            border: none;
            cursor: pointer;
            width: 32px;
            height: 32px;
            padding: 4px;
            border-radius: 4px;
        `;

        waffleButton.onmouseover = () => waffleButton.style.backgroundColor = 'var(--colorNeutralBackground1Hover, #f0f0f0)';
        waffleButton.onmouseout = () => waffleButton.style.backgroundColor = 'transparent';

        navHeader.prepend(waffleButton);
        waffleButton.onclick = openWaffleModal;
    }

    function openWaffleModal() {
        // Create the main overlay
        const modalContainer = document.createElement('div');
        modalContainer.id = 'waffle-modal-container';
        modalContainer.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; z-index:10000; background:rgba(0,0,0,0.65); display:flex; align-items:center; justify-content:center;';
        
        // Create the loader element
        const loader = document.createElement('div');
        loader.id = 'waffle-loader';
        loader.style.cssText = 'display:flex; flex-direction:column; align-items:center; gap:20px; color:white; text-shadow:0 1px 2px rgba(0,0,0,.5); transition:opacity .3s ease-out; animation:waffleFadeIn .3s;';
        loader.innerHTML = '<div class="waffle-spinner"></div><div>Loading M365 Apps...</div>';

        // Create the main dialog window (starts invisible)
        const modalDialog = document.createElement('div');
        modalDialog.style.cssText = 'width:min(95vw, 999px); height:90vh; border-radius:8px; box-shadow:0 5px 20px rgba(0,0,0,.3); overflow:hidden; background-color:' + window.getComputedStyle(document.body).backgroundColor + '; opacity:0; transition:opacity .4s ease-in; animation:waffleFadeInUp .4s ease-out;';
        
        const iframe = document.createElement('iframe');
        iframe.src = 'https://m365.cloud.microsoft/apps/';
        iframe.style.cssText = 'width:100%; height:100%; border:none;';

        const closeBtn = document.createElement('button');
        closeBtn.innerHTML = '×';
        closeBtn.style.cssText = 'position:absolute; top:2vh; right:2vw; z-index:10001; font-size:28px; font-weight:bold; color:white; background:rgba(0,0,0,.4); border:none; cursor:pointer; border-radius:50%; width:36px; height:36px; line-height:36px; opacity:0; transition:opacity .3s, background-color .2s;';
        closeBtn.onmouseover = () => closeBtn.style.background = 'rgba(0,0,0,0.7)';
        closeBtn.onmouseout = () => closeBtn.style.background = 'rgba(0,0,0,0.4)';
        
        const closeModal = () => {
            document.removeEventListener('keydown', handleEsc);
            modalContainer.remove();
        };
        const handleEsc = (event) => { if (event.key === 'Escape') closeModal(); };

        closeBtn.onclick = closeModal;
        modalContainer.onclick = (event) => { if (event.target === modalContainer) closeModal(); };
        document.addEventListener('keydown', handleEsc);
        
        iframe.onload = () => {
            // Reveal content
            setTimeout(() => {
                loader.style.opacity = '0';
                modalDialog.style.opacity = '1';
                closeBtn.style.opacity = '1';
                setTimeout(() => loader.remove(), 300);
            }, 200); // A small delay for a smooth visual transition

            // Aggressively clean the iframe DOM
            const iDoc = iframe.contentDocument;
            if (iDoc) {
                const cleanupInterval = setInterval(() => {
                    const sidebar = iDoc.querySelector('div[role="navigation"]');
                    const helpWidget = iDoc.querySelector('div[id*="Microsoft_Webchat_"]');
                    const header = iDoc.querySelector('#O365_NavHeader');
                    if (sidebar) sidebar.remove();
                    if (helpWidget) helpWidget.remove();
                    if (header) header.remove();
                }, 50);
                setTimeout(() => clearInterval(cleanupInterval), 2000); // Stop after 2 seconds
            }
        };

        modalDialog.appendChild(iframe);
        modalContainer.appendChild(loader);
        modalContainer.appendChild(modalDialog);
        modalContainer.appendChild(closeBtn);
        document.body.appendChild(modalContainer);
    }
})();