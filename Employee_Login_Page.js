/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/ui/serverWidget', 'N/search', 'N/log', 'N/crypto', 'N/record', 'N/runtime'],
function (ui, search, log, crypto, record, runtime) {

  function onRequest(context) {
    const isGET = context.request.method === 'GET';
    const params = context.request.parameters || {};
    const view = (params.view || '').toLowerCase(); // '' | 'profile'
    const form = ui.createForm({ title: 'Molisana Portal' }); // header hidden via CSS

    // Keep your extforms URL so links retain script/deploy/compid/ns-at
    const currentUrl = "https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2110&deploy=1&compid=4975346&ns-at=AAEJ7tMQamzukv1WMqTK6i2c27bRetbrd2MDLjhDgPPFOawMxCo";
    const joiner = currentUrl.indexOf('?') > -1 ? '&' : '?';

    // --- Signed session helpers (so GET can be "logged in") ---
    const SECRET = runtime.getCurrentScript().getParameter({ name: 'custscript_portal_secret' }) || 'change-me';
    const TOKEN_TTL_MS = 30 * 60 * 1000; // 30 minutes

    function sign(empid, ts) {
      const h = crypto.createHash({ algorithm: crypto.HashAlg.SHA256 });
      h.update({ input: empid + '|' + ts + '|' + SECRET });
      return h.digest({ outputEncoding: crypto.Encoding.HEX });
    }
    function verify(empid, ts, sig) {
      if (!empid || !ts || !sig) return false;
      if (Math.abs(Date.now() - parseInt(ts, 10)) > TOKEN_TTL_MS) return false;
      try { return sign(empid, ts) === sig; } catch (e) { log.error('verify token', e); return false; }
    }

    // --- Base CSS/HTML (needed before any branch that appends) ---
    let html = `
      <style>
        #div__title, .uir-page-title, .uir-record-type, .uir-page-header, .ns-main-form-title { display:none !important; }
        html, body { height:100%; }
        body { background:#000; margin:0; font-family:-apple-system, BlinkMacSystemFont,'Segoe UI',Roboto,Oxygen; color:#111; }

        .login-container, .profile-container{
          width:100%; max-width:420px; margin:100px auto; padding:28px 24px;
          background:#fff; border-radius:14px; box-shadow:0 10px 28px rgba(0,0,0,.45);
        }
        .login-container img.logo, .profile-container img.logo{ display:block; margin:0 auto 14px; width:120px; height:auto; }
        h2{ font-size:22px; margin:0 0 16px; color:#111; text-align:center; }
        label{ display:block; font-size:13px; color:#374151; margin:10px 0 6px; font-weight:700; }
        input[type="email"], input[type="password"]{
          width:100%; padding:12px 12px; font-size:16px; border:1px solid #d0d7de; border-radius:10px; background:#fff; box-sizing:border-box;
        }
        .submit-btn{
          width:100%; padding:14px; font-size:16px; font-weight:700;
          background:#0b5cff; color:#fff; border:none; border-radius:10px; cursor:pointer;
          transition:transform .08s ease, box-shadow .08s ease, background .2s ease;
          box-shadow:0 6px 16px rgba(11,92,255,.25); margin-top:14px;
        }
        .submit-btn:hover{ background:#094bcc; transform:translateY(-1px); box-shadow:0 10px 22px rgba(11,92,255,.28); }
        .link-btn{ display:inline-block; margin-top:12px; text-decoration:none; color:#0b5cff; font-weight:700; text-align:center; width:100%; }
        .msg{ margin-top:12px; font-weight:700; text-align:center; }
        .msg.error{ color:#d32f2f; }
        .msg.success{ color:#2e7d32; }

        .dash-wrap{ max-width:1200px; margin:36px auto 48px; padding:0 18px; }
        .welcome{ display:flex; align-items:center; gap:10px; margin:0 0 16px; color:#e5e7eb; font-weight:700; }
        .welcome .dot{ width:10px; height:10px; border-radius:50%; background:#22c55e; box-shadow:0 0 0 3px rgba(34,197,94,.25); }
        .tiles{ display:grid; grid-template-columns:repeat(auto-fill,minmax(260px,1fr)); gap:18px; }
        .tile{ background:#fff; border:1px solid #e5e7eb; border-radius:16px; overflow:hidden; box-shadow:0 8px 24px rgba(0,0,0,.25); transition:transform .12s ease, box-shadow .12s ease; }
        .tile:hover{ transform:translateY(-3px); box-shadow:0 14px 34px rgba(0,0,0,.32); }
        .tile-link{ display:block; text-decoration:none; color:inherit; }
        .tile-hero{ height:120px; display:flex; align-items:center; justify-content:center; background:linear-gradient(135deg,#eef2ff 0%,#f8fafc 100%); }
        .tile-hero.inv{ background:linear-gradient(135deg,#e0f2fe 0%,#f0f9ff 100%); }
        .tile-hero.price{ background:linear-gradient(135deg,#fee2e2 0%,#fff1f2 100%); }
        .tile-hero.reorder{ background:linear-gradient(135deg,#dcfce7 0%,#f0fdf4 100%); }
        .tile-hero.po{ background:linear-gradient(135deg,#ede9fe 0%,#f5f3ff 100%); }
        .tile-hero.profile{ background:linear-gradient(135deg,#fef9c3 0%,#fffbeb 100%); }
        .tile-hero svg{ width:84px; height:84px; color:#0b5cff; }
        .tile-body{ padding:14px 16px 16px; }
        .tile-title{ font-size:16px; font-weight:800; margin:0 0 6px; color:#0f172a; display:flex; align-items:center; gap:8px; }
        .badge{ font-size:11px; padding:2px 8px; border-radius:999px; background:#f1f5f9; color:#334155; font-weight:700; }
        .tile-desc{ font-size:13px; color:#475569; margin:0; line-height:1.45; }
        .no-access{ color:#e5e7eb; text-align:center; margin-top:28px; }
      </style>
    `;

    // ---------- Signed GET -> show dashboard without re-login ----------
    if (isGET && view !== 'profile' && params.empid && params.ts && params.sig && verify(params.empid, params.ts, params.sig)) {
      const loggedInId = params.empid;
      const lf = search.lookupFields({
        type: search.Type.EMPLOYEE,
        id: parseInt(loggedInId, 10),
        columns: ['custentity_mi_price_level', 'custentity_external_portal_access']
      });
      const priceLevelbase = (lf.custentity_mi_price_level || []);
      var priceLevel = [];

      for (let index = 0; index < priceLevelbase.length; index++) {
        priceLevel.push(parseInt(priceLevelbase[index].value))
      }
      log.debug('priceLevel', priceLevel)

      const accessSet = new Set();
      const acc = lf.custentity_external_portal_access;
      if (Array.isArray(acc)) {
        acc.forEach(v => accessSet.add(String((v && v.value) || v)));
      } else if (acc) {
        String(acc).split(',').map(s => s.trim()).filter(Boolean).forEach(v => accessSet.add(v));
      }

      const ts = params.ts;
      const sig = params.sig;

      // URLs for tiles
      const urlItemInventory = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2107&deploy=1&compid=4975346&ns-at=AAEJ7tMQCF0_DxXiMjqflLLzgpO9FsliNEe7-SwDzKCWE2BzC-4&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
      const urlPriceChange   = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2311&deploy=1&compid=4975346&ns-at=AAEJ7tMQL9UOA94WkZJa9hx-tE5vj4RGzOoM-JVdXqNbMHypnes&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
      const urlMIReorder     = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2108&deploy=1&compid=4975346&ns-at=AAEJ7tMQmJxVsovhMpsEMUF39xnBuyMwWM4G2T7SnvA62twq8hg&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
      const urlPlannedPO     = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2313&deploy=1&compid=4975346&ns-at=AAEJ7tMQw0WmH41Gs02OQcTCXn0VTpmqW6ShNtGIxBftDD0rD7c&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
      const urlAvailTool     = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2644&deploy=1&compid=4975346&ns-at=AAEJ7tMQZ1fX7Js6L66rPn3Jy3YA_j6ntOznzRLSRFWZPZhUQCc&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;

      const svgInv = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><rect x="3" y="4" width="18" height="14" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M3 9h18M8 13h4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
      const svgPrice = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 12a8 8 0 1 0 16 0" stroke="currentColor" stroke-width="1.5"/><path d="M12 6v12M9 9h6M9 15h6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
      const svgReorder = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 7h16M4 12h10M4 17h7" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><circle cx="18" cy="12" r="3" stroke="currentColor" stroke-width="1.5"/></svg>`;
      const svgPO = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M8 7h8M8 11h8M8 15h5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
      const svgProfile = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><circle cx="12" cy="8" r="4" stroke="currentColor" stroke-width="1.5"/><path d="M4 20c1.5-4 6.5-4 8-4s6.5 0 8 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
      const svgAvail = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
  <rect x="3" y="4" width="18" height="14" rx="2" stroke="currentColor" stroke-width="1.5"/>
  <path d="M7 9h4M7 13h3M7 17h6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
  <circle cx="16.5" cy="12" r="2.25" stroke="currentColor" stroke-width="1.5"/>
</svg>`;

      
      const has = (id) => accessSet.has(String(id));
      const hasAny = (ids) => ids.some(id => accessSet.has(String(id)));

      let tilesHtml = '';
      if (has(1)) { tilesHtml += `
        <div class="tile">
          <a class="tile-link" href="${urlItemInventory}" target="_blank" rel="noopener">
            <div class="tile-hero inv">${svgInv}</div>
            <div class="tile-body"><p class="tile-title">Item Inventory Dashboard <span class="badge">Personalized</span></p><p class="tile-desc">Live inventory & availability (uses your ID and price level).</p></div>
          </a>
        </div>`; }
      if (has(2)) { tilesHtml += `
        <div class="tile">
          <a class="tile-link" href="${urlPriceChange}" target="_blank" rel="noopener">
            <div class="tile-hero price">${svgPrice}</div>
            <div class="tile-body"><p class="tile-title">Price Change Tool</p><p class="tile-desc">Propose and review item price updates in a guided workflow.</p></div>
          </a>
        </div>`; }
      if (hasAny([3,4])) { tilesHtml += `
        <div class="tile">
          <a class="tile-link" href="${urlMIReorder}" target="_blank" rel="noopener">
            <div class="tile-hero reorder">${svgReorder}</div>
            <div class="tile-body"><p class="tile-title">MI Reorder Tool</p><p class="tile-desc">Analyze reorder points and create replenishment plans.</p></div>
          </a>
        </div>`; }
      if (hasAny([5,6])) { tilesHtml += `
        <div class="tile">
          <a class="tile-link" href="${urlPlannedPO}" target="_blank" rel="noopener">
            <div class="tile-hero po">${svgPO}</div>
            <div class="tile-body"><p class="tile-title">Planned PO Approval</p><p class="tile-desc">Review and approve planned purchase orders.</p></div>
          </a>
        </div>`; }
     if (hasAny([5,6])) { tilesHtml += `
        <div class="tile">
          <a class="tile-link" href="${urlPlannedPO}" target="_blank" rel="noopener">
            <div class="tile-hero po">${svgPO}</div>
            <div class="tile-body"><p class="tile-title">Planned PO Approval</p><p class="tile-desc">Review and approve planned purchase orders.</p></div>
          </a>
        </div>`; }
      if (has(7)) { tilesHtml += `
        <div class="tile">
          <a class="tile-link" href="${urlAvailTool}" target="_blank" rel="noopener">
            <div class="tile-hero price">${svgAvail}</div>
            <div class="tile-body"><p class="tile-title">Item Availability Tool</p><p class="tile-desc">Check available-to-promise, on-order, and commited status by item.</p></div>
          </a>
        </div>`; }

      // Always show Update Profile
      tilesHtml += `
        <div class="tile">
          <a class="tile-link" href="${currentUrl + joiner}view=profile&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}">
            <div class="tile-hero profile">${svgProfile}</div>
            <div class="tile-body"><p class="tile-title">Update Profile</p><p class="tile-desc">Change your portal password.</p></div>
          </a>
        </div>`;

      html += `
        <div class="dash-wrap">
          <div class="welcome"><span class="dot"></span><span>Welcome</span></div>
          <div class="tiles">${tilesHtml}</div>
        </div>`;
      const htmlFieldSigned = form.addField({ id: 'custpage_htmlfield', label: ' ', type: ui.FieldType.INLINEHTML });
      htmlFieldSigned.defaultValue = html;
      context.response.writePage(form);
      return;
    }

    // ---------------- PROFILE (GET/POST) ----------------
    if (view === 'profile') {
      if (isGET) {
        const empid = params.empid || '';
        const ts = params.ts || Date.now().toString();
        const sig = params.sig || sign(empid, ts);

        html += `
          <div class="profile-container">
            <img class="logo" src="https://4975346.app.netsuite.com/core/media/media.nl?id=4770&c=4975346&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd" alt="Logo" />
            <h2>Update Profile</h2>
            <form method="POST" id="pwForm">
              <input type="hidden" name="view" value="profile" />
              <input type="hidden" name="action" value="changePw" />
              <input type="hidden" name="empid" value="${empid}" />
              <input type="hidden" name="ts" value="${ts}" />
              <input type="hidden" name="sig" value="${sig}" />
              <label for="oldpw">Old Password</label>
              <input type="password" id="oldpw" name="oldpw" placeholder="Enter current password" required />
              <label for="newpw">New Password</label>
              <input type="password" id="newpw" name="newpw" placeholder="Enter new password" required />
              <label for="newpw2">Re-enter New Password</label>
              <input type="password" id="newpw2" name="newpw2" placeholder="Re-enter new password" required />
              <button type="submit" class="submit-btn">Save Changes</button>
              <a class="link-btn" href="${currentUrl + joiner}empid=${encodeURIComponent(empid)}&ts=${ts}&sig=${sig}">← Back to Dashboard</a>
            </form>
          </div>
          <script>
            (function(){
              var f = document.getElementById('pwForm');
              f.addEventListener('submit', function(ev){
                var oldpw = f.oldpw.value.trim();
                var np = f.newpw.value.trim();
                var np2 = f.newpw2.value.trim();
                if(!oldpw || !np || !np2){ alert('Please fill out all fields.'); ev.preventDefault(); return false; }
                if(np !== np2){ alert('New passwords do not match.'); ev.preventDefault(); return false; }
                if(!confirm('Are you sure you want to change your password?')){ ev.preventDefault(); return false; }
                return true;
              });
            })();
          </script>
        `;
      } else {
        // POST: enforce validation ORDER server-side
        const action = (params.action || '').toLowerCase();
        if (action === 'changepw') {
          const empid = params.empid;
          const ts = params.ts || Date.now().toString();
          const sig = params.sig || sign(empid, ts);

          const oldpw = (params.oldpw || '').trim();
          const newpw = (params.newpw || '').trim();
          const newpw2 = (params.newpw2 || '').trim();

          let err = '';

          // 1) all fields present
          if (!empid || !oldpw || !newpw || !newpw2) {
            err = 'Please fill out all fields.';
          }
          // 2) new passwords match
          else if (newpw !== newpw2) {
            err = 'New passwords do not match.';
          }
          // 3) old password matches employee profile
          else {
            try {
              const ok = crypto.checkPasswordField({
                value: oldpw,
                recordType: record.Type.EMPLOYEE,
                recordId: parseInt(empid, 10),
                fieldId: 'custentity_external_portal_password'
              });
              if (!ok) err = 'Old password is incorrect.';
            } catch (e) {
              log.error('checkPasswordField error', e);
              err = 'Password validation error.';
            }
          }

          if (err) {
            html += `
              <div class="profile-container">
                <img class="logo" src="https://4975346.app.netsuite.com/core/media/media.nl?id=4770&c=4975346&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd" alt="Logo" />
                <h2>Update Profile</h2>
                <div class="msg error">${err}</div>
                <form method="POST" id="pwForm">
                  <input type="hidden" name="view" value="profile" />
                  <input type="hidden" name="action" value="changePw" />
                  <input type="hidden" name="empid" value="${empid || ''}" />
                  <input type="hidden" name="ts" value="${ts}" />
                  <input type="hidden" name="sig" value="${sig}" />
                  <label for="oldpw">Old Password</label>
                  <input type="password" id="oldpw" name="oldpw" placeholder="Enter current password" required />
                  <label for="newpw">New Password</label>
                  <input type="password" id="newpw" name="newpw" placeholder="Enter new password" required />
                  <label for="newpw2">Re-enter New Password</label>
                  <input type="password" id="newpw2" name="newpw2" placeholder="Re-enter new password" required />
                  <button type="submit" class="submit-btn">Save Changes</button>
                  <a class="link-btn" href="${currentUrl + joiner}empid=${encodeURIComponent(empid)}&ts=${ts}&sig=${sig}">← Back to Dashboard</a>
                </form>
              </div>
              <script>
                (function(){
                  var f = document.getElementById('pwForm');
                  f.addEventListener('submit', function(ev){
                    var oldpw = f.oldpw.value.trim(); var np = f.newpw.value.trim(); var np2 = f.newpw2.value.trim();
                    if(!oldpw || !np || !np2){ alert('Please fill out all fields.'); ev.preventDefault(); return false; }
                    if(np !== np2){ alert('New passwords do not match.'); ev.preventDefault(); return false; }
                    if(!confirm('Are you sure you want to change your password?')){ ev.preventDefault(); return false; }
                    return true;
                  });
                })();
              </script>
            `;
          } else {
            // 4) update only after all validations pass
            try {
              record.submitFields({
                type: record.Type.EMPLOYEE,
                id: parseInt(empid, 10),
                values: { custentity_external_portal_password: newpw }
              });
              html += `
                <div class="profile-container">
                  <img class="logo" src="https://4975346.app.netsuite.com/core/media/media.nl?id=4770&c=4975346&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd" alt="Logo" />
                  <h2>Update Profile</h2>
                  <div class="msg success">Password updated successfully.</div>
                  <a class="link-btn" href="${currentUrl + joiner}empid=${encodeURIComponent(empid)}&ts=${ts}&sig=${sig}">← Back to Dashboard</a>
                </div>
              `;
            } catch (e) {
              log.error('submitFields password update error', e);
              html += `
                <div class="profile-container">
                  <img class="logo" src="https://4975346.app.netsuite.com/core/media/media.nl?id=4770&c=4975346&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd" alt="Logo" />
                  <h2>Update Profile</h2>
                  <div class="msg error">Failed to update password.</div>
                  <a class="link-btn" href="${currentUrl + joiner}empid=${encodeURIComponent(empid)}&ts=${ts}&sig=${sig}">← Back to Dashboard</a>
                </div>
              `;
            }
          }
        }

        const profField = form.addField({ id: 'custpage_htmlfield', label: ' ', type: ui.FieldType.INLINEHTML });
        profField.defaultValue = html;
        context.response.writePage(form);
        return; // stop here for profile view
      }
    }

    // ---------------- LOGIN / DASHBOARD (POST) ----------------
    let message = '';
    let messageClass = 'msg';

    if (!isGET) {
      const email = (params.email || '').trim();
      const password = (params.password || '').trim();
      log.debug('Login Attempt', { params, email, password });

      let isValid = false;
      let loggedInId = '';
      let priceLevel = '';
      const accessSet = new Set();

      const empSearch = search.create({
        type: search.Type.EMPLOYEE,
        filters: [['email','is',email]],
        columns: ['internalid', 'custentity_mi_price_level', 'custentity_external_portal_access']
      });

      empSearch.run().each(r => {
        const empId = r.getValue('internalid');
        log.debug('empId', {empId})
        try {
          const ok = crypto.checkPasswordField({
            value: password,
            recordType: record.Type.EMPLOYEE,
            recordId: parseInt(empId, 10),
            fieldId: 'custentity_external_portal_password'
          });
          if (ok) {
            loggedInId = empId;
            priceLevel = r.getValue('custentity_mi_price_level') || '';
            const accessVal = r.getValue('custentity_external_portal_access');
            if (Array.isArray(accessVal)) accessVal.forEach(v => accessSet.add(String(v)));
            else if (accessVal) String(accessVal).split(',').map(s => s.trim()).filter(Boolean).forEach(v => accessSet.add(v));
            isValid = true;
            return false;
          }
        } catch (e) { log.error('checkPasswordField error', e); }
        return true;
      });

      if (isValid) {
        message = 'Logged in';
        messageClass = 'msg success';

        const ts = Date.now().toString();
        const sig = sign(loggedInId, ts);

        const has = (id) => accessSet.has(String(id));
        const hasAny = (ids) => ids.some(id => accessSet.has(String(id)));


        const svgAvail = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
  <rect x="3" y="4" width="18" height="14" rx="2" stroke="currentColor" stroke-width="1.5"/>
  <path d="M7 9h4M7 13h3M7 17h6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
  <circle cx="16.5" cy="12" r="2.25" stroke="currentColor" stroke-width="1.5"/>
</svg>`;

        

        const urlItemInventory  = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2107&deploy=1&compid=4975346&ns-at=AAEJ7tMQCF0_DxXiMjqflLLzgpO9FsliNEe7-SwDzKCWE2BzC-4&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
        const urlPriceChange    = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2311&deploy=1&compid=4975346&ns-at=AAEJ7tMQL9UOA94WkZJa9hx-tE5vj4RGzOoM-JVdXqNbMHypnes&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
        const urlMIReorder3     = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2108&deploy=1&compid=4975346&ns-at=AAEJ7tMQmJxVsovhMpsEMUF39xnBuyMwWM4G2T7SnvA62twq8hg&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
        const urlMIReorder4     = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2108&deploy=1&compid=4975346&ns-at=AAEJ7tMQmJxVsovhMpsEMUF39xnBuyMwWM4G2T7SnvA62twq8hg&type=4&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
        const urlPlannedPO5     = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2313&deploy=1&compid=4975346&ns-at=AAEJ7tMQw0WmH41Gs02OQcTCXn0VTpmqW6ShNtGIxBftDD0rD7c&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
        const urlPlannedPO6     = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2520&deploy=1&compid=4975346&ns-at=AAEJ7tMQVziJF1qjHHrnbuPycJ3uz76kG7ifpvym_jIq-Whr50U&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;
        const urlAvailTool      = `https://4975346.extforms.netsuite.com/app/site/hosting/scriptlet.nl?script=2644&deploy=1&compid=4975346&ns-at=AAEJ7tMQZ1fX7Js6L66rPn3Jy3YA_j6ntOznzRLSRFWZPZhUQCc&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}`;

        const svgInv = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><rect x="3" y="4" width="18" height="14" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M3 9h18M8 13h4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
        const svgPrice = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 12a8 8 0 1 0 16 0" stroke="currentColor" stroke-width="1.5"/><path d="M12 6v12M9 9h6M9 15h6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
        const svgReorder = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 7h16M4 12h10M4 17h7" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><circle cx="18" cy="12" r="3" stroke="currentColor" stroke-width="1.5"/></svg>`;
        const svgPO = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M8 7h8M8 11h8M8 15h5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
        const svgProfile = `<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><circle cx="12" cy="8" r="4" stroke="currentColor" stroke-width="1.5"/><path d="M4 20c1.5-4 6.5-4 8-4s6.5 0 8 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;

        let tilesHtml = '';
        if (has(1)) tilesHtml += `
          <div class="tile"><a class="tile-link" href="${urlItemInventory}" target="_blank" rel="noopener"><div class="tile-hero inv">${svgInv}</div><div class="tile-body"><p class="tile-title">Item Inventory Dashboard <span class="badge">Personalized</span></p><p class="tile-desc">Live inventory & availability (uses your ID and price level).</p></div></a></div>`;
        if (has(2)) tilesHtml += `
          <div class="tile"><a class="tile-link" href="${urlPriceChange}" target="_blank" rel="noopener"><div class="tile-hero price">${svgPrice}</div><div class="tile-body"><p class="tile-title">Price Change Tool</p><p class="tile-desc">Propose and review item price updates in a guided workflow.</p></div></a></div>`;
        if (has(3)) tilesHtml += `
          <div class="tile"><a class="tile-link" href="${urlMIReorder3}" target="_blank" rel="noopener"><div class="tile-hero reorder">${svgReorder}</div><div class="tile-body"><p class="tile-title">MI Reorder Tool (Admin)</p><p class="tile-desc">Analyze reorder points and create replenishment plans.</p></div></a></div>`;
        if (has(4)) tilesHtml += `
          <div class="tile"><a class="tile-link" href="${urlMIReorder4}" target="_blank" rel="noopener"><div class="tile-hero reorder">${svgReorder}</div><div class="tile-body"><p class="tile-title">MI Reorder Tool (Basic)</p><p class="tile-desc">Analyze reorder points and create replenishment plans.</p></div></a></div>`;
        if (has(5)) tilesHtml += `
          <div class="tile"><a class="tile-link" href="${urlPlannedPO5}" target="_blank" rel="noopener"><div class="tile-hero po">${svgPO}</div><div class="tile-body"><p class="tile-title">Planned PO Approval (Admin)</p><p class="tile-desc">Review and approve planned purchase orders.</p></div></a></div>`;
        if (has(6)) tilesHtml += `
          <div class="tile"><a class="tile-link" href="${urlPlannedPO6}" target="_blank" rel="noopener"><div class="tile-hero po">${svgPO}</div><div class="tile-body"><p class="tile-title">Planned PO Approval (Basic)</p><p class="tile-desc">Review and approve planned purchase orders.</p></div></a></div>`;
        if (has(7)) tilesHtml += `
          <div class="tile"><a class="tile-link" href="${urlAvailTool}" target="_blank" rel="noopener"><div class="tile-hero po">${svgAvail}</div><div class="tile-body"><p class="tile-title">Item Availability Tool</p><p class="tile-desc">Check available-to-promise, on-order, and commited status by item.</p></div></a></div>`;
        tilesHtml += `
          <div class="tile">
            <a class="tile-link" href="${currentUrl + joiner}view=profile&empid=${encodeURIComponent(loggedInId)}&ts=${ts}&sig=${sig}">
              <div class="tile-hero profile">${svgProfile}</div>
              <div class="tile-body"><p class="tile-title">Update Profile</p><p class="tile-desc">Change your portal password.</p></div>
            </a>
          </div>`;

        html += `
          <div class="dash-wrap">
            <div class="welcome"><span class="dot"></span><span>Welcome</span></div>
            <div class="tiles">${tilesHtml}</div>
          </div>
        `;
      } else {
        message = 'Invalid Email or Password';
        messageClass = 'msg error';
      }
    }

    // ---------------- LOGIN (GET) or failed login ----------------
    const showLogin = isGET || messageClass.indexOf('error') !== -1;
    if (showLogin && view !== 'profile') {
      html += `
        <div class="login-container">
          <img class="logo" src="https://4975346.app.netsuite.com/core/media/media.nl?id=4770&c=4975346&h=sGvHCgrcrHMJzoZjmKF-9Og7Y_nEEncIbLigVYoLtPhS1CXd" alt="Logo" />
          <h2>Log In</h2>
          <form method="POST">
            <label for="email">Email</label>
            <input type="email" id="email" name="email" placeholder="Email address" required />
            <label for="password">Password</label>
            <input type="password" id="password" name="password" placeholder="Password" required />
            <button type="submit" class="submit-btn">Login</button>
            ${message ? `<div class="${messageClass}">${message}</div>` : ''}
          </form>
        </div>
      `;
    }

    const htmlField = form.addField({ id: 'custpage_htmlfield', label: ' ', type: ui.FieldType.INLINEHTML });
    htmlField.defaultValue = html;
    context.response.writePage(form);
  }

  return { onRequest };
});
