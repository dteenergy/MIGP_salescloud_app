const cds = require('@sap/cds');
// const { SELECT, INSERT, UPDATE, DELETE } = cds.ql;

const BATCH_SIZE = 200;

// Helper: cast any value to a safe string
const asStr = (v) => (v == null ? '' : String(v));

// ADD this mapping helper before oppCols:
const oPhaseMap = {
  // LCVP
  'Z5': 'Origination',
  'Z6': 'Introduction',
  'Z7': 'Discovery',
  'Z8': 'Structuring',
  'Z0': 'Approval',
  'Z9': 'Contracting',
  'Z3': 'Final Review',
  'Z4': 'Closure',
  // SMB
  'Z1': 'Customer Appointment',
  'Z2': 'Awaiting Confirmation'
};

module.exports = (srv) => {

  // ─── CloudService ──────────────────────────────────────────────────────────
  if (srv.name === 'CloudService') {


    //User Roles
    srv.on('userDetails', async (req) => {
      // BUG 17908: Guard against missing authInfo on transient XSUAA token failures
      if (!req.user || !req.user.authInfo || !req.user.authInfo.token) {
        console.error('userDetails: req.user.authInfo missing — likely transient XSUAA issue');
        return req.error(401, 'Authorization token unavailable. Please retry.');
      }

      const roles = Object.keys(req.user.roles || {});
      console.log(roles);
      let userName = req.user.authInfo.token.getPayload().user_name, //U id if DTE SSO otherwise email
        email = req.user.authInfo.token.getPayload().email,
        hasUserAccess = false,
        hasAdminAccess = false;

      for (var i = 0; i < roles.length; i++) {
        let scope = roles[i];
        switch (scope) {
          // BUG 17908: XSUAA delivers scopes as "xsappname.ScopeName" in production
          // Match both bare name (local dev) and prefixed name (BTP production)
          case "MIGP_User":
          case "salescloud.MIGP_User":
            hasUserAccess = true;
            break;
          case "MIGP_Admin":
          case "salescloud.MIGP_Admin":
            hasAdminAccess = true;
            break;
          default:
          // code block
        }
      }
      let userInfo = {
        "userName": userName,
        "email": email,
        "hasUserAccess": hasUserAccess,
        "hasAdminAccess": hasAdminAccess
      };
      console.log(userInfo);
      return userInfo;

    });


    // ── getOpportunityRecord ────────────────────────────────────────────────
    // FIX: return { value: JSON.stringify(result) } so controller JSON.parse works
    // consistently with getLatestCPIRecord response shape
    srv.on('getOpportunityRecord', async (req) => {
      const oppId = req.data.connection_object;
      console.log('getOpportunityRecord called with oppId:', oppId);

      if (!oppId || !oppId.trim()) {
        return JSON.stringify({ value: 'error', message: 'Opportunity ID is required' });
      }

      try {
        const db = await cds.connect.to('db');
        const { Opportunity, InvolvedParties, ProviderContract, ConsumptionDetail, Prospect } =
          db.entities('com.salescloud');

        const opp = await db.run(
          SELECT.one.from(Opportunity)
            .where({ Oppid: oppId.trim(), objectStatus: 'OPP' })
        );

        if (!opp) {
          console.warn('No opportunity found for Oppid:', oppId);
          return JSON.stringify({ value: 'error', message: 'Not found' });
        }

        // Fetch all 4 child tables in parallel using the UUID primary key
        const [involvedParties, providerContracts, consumptionDetails, prospects] =
          await Promise.all([
            db.run(SELECT.from(InvolvedParties).where({ opportunity_ID: opp.ID })),
            db.run(SELECT.from(ProviderContract).where({ opportunity_ID: opp.ID })),
            db.run(SELECT.from(ConsumptionDetail).where({ opportunity_ID: opp.ID })),
            db.run(SELECT.from(Prospect).where({ opportunity_ID: opp.ID })),
          ]);

        console.log(`Fetched children for ${oppId}: IP=${involvedParties.length} PC=${providerContracts.length} CD=${consumptionDetails.length} PR=${prospects.length}`);

        const result = {
          ID: opp.ID,
          Oppid: opp.Oppid,
          quoteId: opp.quoteId || '',
          objectStatus: opp.objectStatus || '',
          oppCategory: opp.oppCategory || '',
          status: opp.status || '',
          annualGross: opp.annualGross || '',
          sizeofOpp: opp.sizeofOpp || '',
          percentUsage: opp.percentUsage || '85',
          annualSubs: opp.annualSubs || '',
          oppSubs: opp.oppSubs || '',
          contLifetimeval: opp.contLifetimeval || '',
          salesType: opp.salesType || '',
          priceStructure: opp.priceStructure || '',
          nte: opp.nte || '',
          annualMWh: opp.annualMWh || '',
          term: opp.term || '',
          meteredCon: opp.meteredCon || '',
          unmeteredCon: opp.unmeteredCon || '',
          commencementLetterSent: opp.commencementLetterSent || '',
          commencementLetterSigned: opp.commencementLetterSigned || '',
          enrolled: opp.enrolled || '',
          includeInPipeline: opp.includeInPipeline || '',
          salesPhase: opp.salesPhase || '',  // ✅ ADD
          enrollmentReferralCode: opp.enrollmentReferralCode || '',   // ← add
          enrollmentEmailId: opp.enrollmentEmailId || '',             // ← add
          lastFetchedAt: opp.lastFetchedAt || null,                   // ← timestamp


          // ── Child tables — map every field explicitly so nothing is lost ──
          involvedParties: involvedParties.map(r => ({
            buspartner: r.buspartner || '',
            role: r.role || ''
          })),

          providerContracts: providerContracts.map(r => ({
            product: r.product || '',
            portfolio: r.portfolio || '',
            fixedMWh: r.fixedMWh || '',
            fixedPrice: r.fixedPrice || '',
            portfolioprice: r.portfolioprice || '',
            subspercent: r.subspercent || '',
            startDate: r.startDate || '',
            endDate: r.endDate || '',
            estUsage: r.estUsage || '',
            recPrice: r.recPrice || '',
            migpPercent: r.migpPercent || '',
            recpercent: r.recpercent || '',
            migpMWh: r.migpMWh || '',
            recMWh: r.recMWh || '',
            enrollChk: r.enrollChk === true || r.enrollChk === 'true',  // ✅ boolean
            exportDate: r.exportDate || '',
            exportChk: r.exportChk === true || r.exportChk === 'true',  // ✅ boolean
            netPremium: r.netPremium || '',   // ← ADD
          })),

          consumptionDetails: consumptionDetails.map(r => ({
            contacc: r.contacc || '',
            buspartner: r.buspartner || '',
            validity: r.validity || '',
            replacementCA: r.replacementCA || '',
            est12month: r.est12month || '',
            providercontractId: r.providercontractId || '',
            cycle20: r.cycle20 || '',
            metered: r.metered || '',
            arrear60: r.arrear60 || '',
            nsfFlag: r.nsfFlag || '',
            inactive: r.inactive || '',
            enrolled: r.enrolled || '',
            // FIX: selected stored as string "true"/"false" — convert back to boolean
            selected: r.selected === 'true' || r.selected === true
          })),

          prospects: prospects.map(r => ({
            siteAddLoc: r.siteAddLoc || '',
            projectedCon: r.projectedCon || '',
            year: r.year || ''
          }))
        };
        console.log('lastFetchedAt from DB:', opp.lastFetchedAt);

        // FIX: wrap inner result as JSON string so controller unwrap logic is consistent
        return JSON.stringify({ value: JSON.stringify(result) });

      } catch (e) {
        console.error('getOpportunityRecord error:', e.message);
        return JSON.stringify({ value: 'error', message: e.message });
      }
    });

    // ── saveOpportunity (header-only save) ──────────────────────────────────
    srv.on('saveOpportunity', async (req) => {
      const { opp } = req.data;
      if (!opp || !opp.Oppid || !opp.Oppid.trim()) {
        return JSON.stringify({ value: 'error', message: 'Opportunity ID is required' });
      }

      try {
        const db = await cds.connect.to('db');
        const { Opportunity } = db.entities('com.salescloud');

        const isQuote = (opp.objectStatus === 'QUOTE');
        const existing = await db.run(
          SELECT.one.from(Opportunity)
            .columns('ID')
            .where(isQuote
              ? { Oppid: opp.Oppid.trim(), objectStatus: 'QUOTE', quoteId: opp.quoteId }
              : { Oppid: opp.Oppid.trim(), objectStatus: 'OPP' }
            )
        );

        const oppData = {
          quoteId: opp.quoteId || '',
          objectStatus: opp.objectStatus || '',
          oppCategory: opp.oppCategory || '',
          status: opp.status || '',
          annualGross: opp.annualGross || '',
          sizeofOpp: opp.sizeofOpp || '',
          percentUsage: opp.percentUsage || '',
          annualSubs: opp.annualSubs || '',
          oppSubs: opp.oppSubs || '',
          contLifetimeval: opp.contLifetimeval || '',
          salesType: opp.salesType || '',
          priceStructure: opp.priceStructure || '',
          nte: opp.nte || '',
          annualMWh: opp.annualMWh || '',
          term: opp.term || '',
          meteredCon: opp.meteredCon || '',
          unmeteredCon: opp.unmeteredCon || '',
          commencementLetterSent: opp.commencementLetterSent || '',
          commencementLetterSigned: opp.commencementLetterSigned || '',
          enrolled: opp.enrolled || '',
          includeInPipeline: opp.includeInPipeline || '',
          salesPhase: opp.salesPhase || '',
          enrollmentReferralCode: opp.enrollmentReferralCode || '',   // ← add
          enrollmentEmailId: opp.enrollmentEmailId || '',             // ← add
          // lastFetchedAt: opp.lastFetchedAt || '',                   // ← timestamp
          // Only persist if it's a valid ISO timestamp — ignore display-formatted strings
          // lastFetchedAt: header.lastFetchedAt && header.lastFetchedAt.includes('T')
          //   ? header.lastFetchedAt
          //   : undefined,
          lastFetchedAt: opp.lastFetchedAt && opp.lastFetchedAt.includes('T')
            ? opp.lastFetchedAt
            : undefined,
        };

        if (existing) {
          await db.run(UPDATE(Opportunity).set(oppData).where({ ID: existing.ID }));
          return JSON.stringify({ value: 'updated', message: 'Opportunity updated successfully' });
        } else {
          const newID = cds.utils.uuid();
          await db.run(
            INSERT.into(Opportunity).entries({ ID: newID, Oppid: opp.Oppid.trim(), ...oppData })
          );
          return JSON.stringify({ value: 'inserted', message: 'Opportunity saved successfully' });
        }

      } catch (e) {
        console.error('saveOpportunity error:', e.message);
        return JSON.stringify({ value: 'error', message: e.message });
      }
    });

    // ── saveFullOpportunity (header + all 4 child tables) ──────────────────
    srv.on('saveFullOpportunity', async (req) => {
      try {
        const db = await cds.connect.to('db');
        const { Opportunity, InvolvedParties, ProviderContract, ConsumptionDetail, Prospect } =
          db.entities('com.salescloud');

        const payload = JSON.parse(req.data.bundle || '{}');
        const { header, involvedParties, providerContracts, consumptionDetails, prospects } = payload;

        if (!header || !header.Oppid || !header.Oppid.trim()) {
          return JSON.stringify({ value: 'error', message: 'Oppid is required' });
        }
        const Oppid = header.Oppid.trim();

        console.log(`saveFullOpportunity: Oppid=${Oppid} IP=${(involvedParties || []).length} PC=${(providerContracts || []).length} CD=${(consumptionDetails || []).length} PR=${(prospects || []).length}`);

        const tx = db.tx(req);

        // 1) Upsert header
        const objectStatus = header.objectStatus || '';
        if (!objectStatus) {
          console.error('saveFullOpportunity: objectStatus missing in header — cannot determine OPP vs QUOTE');
          return JSON.stringify({ value: 'error', message: 'objectStatus is required' });
        }
        const isQuote = (objectStatus === 'QUOTE');
        const existing = await tx.run(
          SELECT.one.from(Opportunity)
            .columns('ID')
            .where(isQuote
              ? { Oppid, objectStatus: 'QUOTE', quoteId: header.quoteId }
              : { Oppid, objectStatus: 'OPP' }
            )
        );

        const oppCols = {
          quoteId: header.quoteId ?? '',
          objectStatus: objectStatus,
          oppCategory: header.oppCategory ?? '',
          status: header.status ?? '',
          annualGross: header.annualGross ?? '',
          sizeofOpp: header.sizeofOpp ?? '',
          percentUsage: header.percentUsage ?? '',
          annualSubs: header.annualSubs ?? '',
          oppSubs: header.oppSubs ?? '',
          contLifetimeval: header.contLifetimeval ?? '',
          salesType: header.salesType ?? '',
          priceStructure: header.priceStructure ?? '',
          nte: header.nte ?? '',
          annualMWh: header.annualMWh ?? '',
          term: header.term ?? '',
          meteredCon: header.meteredCon ?? '',
          unmeteredCon: header.unmeteredCon ?? '',
          commencementLetterSent: header.commencementLetterSent ?? '',
          commencementLetterSigned: header.commencementLetterSigned ?? '',
          enrolled: header.enrolled ?? '',
          includeInPipeline: header.includeInPipeline ?? '',
          salesPhase: header.salesPhase ?? '',  // ✅ ADD
          enrollmentReferralCode: header.enrollmentReferralCode ?? '',  // ✅ ADD
          enrollmentEmailId: header.enrollmentEmailId ?? '',  // ✅ ADD
          // lastFetchedAt: header.lastFetchedAt ?? '',                   // ← timestamp
          // Only persist if it's a valid ISO timestamp — ignore display-formatted strings
          lastFetchedAt: header.lastFetchedAt && header.lastFetchedAt.includes('T')
            ? header.lastFetchedAt
            // : undefined,
            : null,   // ← null instead of undefined — forces explicit DB write

        };
        // Move log here to confirm actual value going into UPDATE
        console.log('lastFetchedAt being saved:', oppCols.lastFetchedAt);

        let oppID;
        if (existing) {
          await tx.run(UPDATE(Opportunity).set(oppCols).where({ ID: existing.ID }));
          oppID = existing.ID;
          console.log('Updated existing opportunity, ID:', oppID);
        } else {
          const newID = cds.utils.uuid();
          await tx.run(INSERT.into(Opportunity).entries({ ID: newID, Oppid, ...oppCols }));
          oppID = newID;
          console.log('Inserted new opportunity, ID:', oppID);
        }

        // 2) Delete all existing children first (replace strategy)
        await tx.run(DELETE.from(InvolvedParties).where({ opportunity_ID: oppID }));
        await tx.run(DELETE.from(ProviderContract).where({ opportunity_ID: oppID }));
        await tx.run(DELETE.from(ConsumptionDetail).where({ opportunity_ID: oppID }));
        await tx.run(DELETE.from(Prospect).where({ opportunity_ID: oppID }));

        // 3) Insert new children
        const ipRows = (involvedParties || []).map(r => ({
          ID: cds.utils.uuid(),
          buspartner: asStr(r.buspartner),
          role: asStr(r.role),
          opportunity_ID: oppID
        }));

        const pcRows = (providerContracts || []).map(r => ({
          ID: cds.utils.uuid(),
          product: asStr(r.product),
          portfolio: asStr(r.portfolio),
          fixedMWh: asStr(r.fixedMWh),
          fixedPrice: asStr(r.fixedPrice),
          portfolioprice: asStr(r.portfolioprice),
          subspercent: asStr(r.subspercent),
          startDate: asStr(r.startDate),
          endDate: asStr(r.endDate),
          estUsage: asStr(r.estUsage),
          recPrice: asStr(r.recPrice),
          migpPercent: asStr(r.migpPercent),
          recpercent: asStr(r.recpercent),
          migpMWh: asStr(r.migpMWh),
          recMWh: asStr(r.recMWh),
          enrollChk: asStr(r.enrollChk),   // ✅ ADD
          exportDate: asStr(r.exportDate),  // ✅ ADD
          exportChk: asStr(r.exportChk),   // ✅ ADD
          netPremium: asStr(r.netPremium),   // ← ADD
          opportunity_ID: oppID
        }));

        const cdRows = (consumptionDetails || []).map(r => ({
          ID: cds.utils.uuid(),
          contacc: asStr(r.contacc),
          buspartner: asStr(r.buspartner),
          validity: asStr(r.validity),
          replacementCA: asStr(r.replacementCA),
          est12month: asStr(r.est12month),
          est12monthMetered: asStr(r.est12monthMetered),    // ← ADD
          est12monthUnMetered: asStr(r.est12monthUnMetered),  // ← ADD
          RECBlender_flag: asStr(r.RECBlender_flag),      // ← ADD
          providercontractId: asStr(r.providercontractId),
          cycle20: asStr(r.cycle20),
          metered: asStr(r.metered),
          arrear60: asStr(r.arrear60),
          nsfFlag: asStr(r.nsfFlag),
          inactive: asStr(r.inactive),
          enrolled: asStr(r.enrolled),
          // FIX: store boolean as string "true"/"false" consistently
          selected: r.selected === true || r.selected === 'true' ? 'true' : 'false',
          opportunity_ID: oppID
        }));

        const prRows = (prospects || []).map(r => ({
          ID: cds.utils.uuid(),
          siteAddLoc: asStr(r.siteAddLoc),
          projectedCon: asStr(r.projectedCon),
          year: asStr(r.year),
          opportunity_ID: oppID
        }));

        // Insert each child table — only if rows exist
        if (ipRows.length) await tx.run(INSERT.into(InvolvedParties).entries(ipRows));
        if (pcRows.length) await tx.run(INSERT.into(ProviderContract).entries(pcRows));
        if (cdRows.length) await tx.run(INSERT.into(ConsumptionDetail).entries(cdRows));
        if (prRows.length) await tx.run(INSERT.into(Prospect).entries(prRows));

        console.log(`saveFullOpportunity committed: IP=${ipRows.length} PC=${pcRows.length} CD=${cdRows.length} PR=${prRows.length}`);
        console.log('lastFetchedAt being saved:', oppCols.lastFetchedAt);
        await tx.commit();
        return JSON.stringify({ value: 'ok' });

      } catch (e) {
        console.error('saveFullOpportunity error:', e.message);
        req.error(500, e.message);
      }
    });

    // ── saveExportDate — silently updates only exportChk + exportDate ──────────
    srv.on('saveExportDate', async (req) => {
      try {
        const db = await cds.connect.to('db');
        const { Opportunity, ProviderContract } = db.entities('com.salescloud');

        const payload = JSON.parse(req.data.bundle || '{}');
        const { oppId, quoteId, objectStatus, providerContracts } = payload;

        if (!oppId || !oppId.trim()) {
          return JSON.stringify({ value: 'error', message: 'oppId is required' });
        }

        const isQuote = (objectStatus === 'QUOTE');
        const existing = await db.run(
          SELECT.one.from(Opportunity).columns('ID')
            .where(isQuote
              ? { Oppid: oppId.trim(), objectStatus: 'QUOTE', quoteId }
              : { Oppid: oppId.trim(), objectStatus: 'OPP' }
            )
        );

        if (!existing) {
          return JSON.stringify({ value: 'error', message: 'Record not found' });
        }

        // Update each PC row individually by product+portfolio key
        for (const r of (providerContracts || [])) {
          await db.run(
            UPDATE(ProviderContract)
              .set({
                exportChk: asStr(r.exportChk),   // 'true' or 'false'
                exportDate: asStr(r.exportDate)   // date string or ''
              })
              .where({
                opportunity_ID: existing.ID,
                product: asStr(r.product),
                portfolio: asStr(r.portfolio)
              })
          );
        }

        console.log(`saveExportDate: updated ${(providerContracts || []).length} PC row(s) for ${isQuote ? 'Quote:' + quoteId : 'Opp:' + oppId}`);
        return JSON.stringify({ value: 'ok' });

      } catch (e) {
        console.error('saveExportDate error:', e.message);
        return JSON.stringify({ value: 'error', message: e.message });
      }
    });


    // ── getLatestCPIRecord ──────────────────────────────────────────────────
    // FIX: renamed cds_detail → cdRows to avoid shadowing the cds module
    srv.on('getLatestCPIRecord', async (req) => {
      const {
        Opportunity,
        InvolvedParties,
        ProviderContract,
        ConsumptionDetail,
        Prospect
      } = cds.entities('com.salescloud');

      try {
        const latest = await cds.run(
          SELECT.one
            .from(Opportunity)
            .columns('ID', 'Oppid', 'quoteId', 'objectStatus', 'oppCategory', 'status', 'salesPhase',
              'annualGross', 'sizeofOpp', 'percentUsage', 'annualSubs',
              'oppSubs', 'contLifetimeval', 'salesType', 'priceStructure',
              'nte', 'annualMWh', 'term', 'meteredCon', 'unmeteredCon',
              'commencementLetterSent', 'commencementLetterSigned',
              'enrolled', 'includeInPipeline', 'modifiedAt', 'enrollmentReferralCode', 'enrollmentEmailId', 'lastFetchedAt')
            .where({ objectStatus: 'OPP' })        // ← ADD
            .orderBy({ modifiedAt: 'desc' })
        );

        if (!latest) {
          return JSON.stringify({ value: 'none', message: 'No CPI records found in database.' });
        }

        const oppID = latest.ID;

        // FIX: renamed cds_detail to cdDetail to avoid shadowing cds module
        const [ips, pcs, cdDetail, prList] = await Promise.all([
          cds.run(SELECT.from(InvolvedParties).where({ opportunity_ID: oppID })),
          cds.run(SELECT.from(ProviderContract).where({ opportunity_ID: oppID })),
          cds.run(SELECT.from(ConsumptionDetail).where({ opportunity_ID: oppID })),
          cds.run(SELECT.from(Prospect).where({ opportunity_ID: oppID }))
        ]);

        const payload = {
          Oppid: latest.Oppid,
          quoteId: latest.quoteId || '',
          objectStatus: latest.objectStatus || '',
          oppCategory: latest.oppCategory || '',
          status: latest.status || '',
          annualGross: latest.annualGross || '',
          sizeofOpp: latest.sizeofOpp || '',
          percentUsage: latest.percentUsage || '85',
          annualSubs: latest.annualSubs || '',
          oppSubs: latest.oppSubs || '',
          contLifetimeval: latest.contLifetimeval || '',
          salesType: latest.salesType || '',
          priceStructure: latest.priceStructure || '',
          nte: latest.nte || '',
          annualMWh: latest.annualMWh || '',
          term: latest.term || '',
          meteredCon: latest.meteredCon || '',
          unmeteredCon: latest.unmeteredCon || '',
          commencementLetterSent: latest.commencementLetterSent || '',
          commencementLetterSigned: latest.commencementLetterSigned || '',
          enrolled: latest.enrolled || '',
          includeInPipeline: latest.includeInPipeline || '',
          salesPhase: latest.salesPhase || '',  // ✅ ADD
          enrollmentReferralCode: latest.enrollmentReferralCode || '',   // ← add
          enrollmentEmailId: latest.enrollmentEmailId || '',             // ← add
          lastFetchedAt: latest.lastFetchedAt || '',
          modifiedAt: latest.modifiedAt,

          involvedParties: ips.map(r => ({
            buspartner: r.buspartner || '',
            role: r.role || ''
          })),

          providerContracts: pcs.map(r => ({
            product: r.product || '',
            portfolio: r.portfolio || '',
            fixedMWh: r.fixedMWh || '',
            fixedPrice: r.fixedPrice || '',
            portfolioprice: r.portfolioprice || '',
            subspercent: r.subspercent || '',
            startDate: r.startDate || '',
            endDate: r.endDate || '',
            estUsage: r.estUsage || '',
            recPrice: r.recPrice || '',
            migpPercent: r.migpPercent || '',
            recpercent: r.recpercent || '',
            migpMWh: r.migpMWh || '',
            recMWh: r.recMWh || '',
            enrollChk: r.enrollChk === true || r.enrollChk === 'true',
            exportDate: r.exportDate || '',  // ✅ ADD
            exportChk: r.exportChk === true || r.exportChk === 'true',
            netPremium: r.netPremium || '',   // ← ADD
          })),

          consumptionDetails: cdDetail.map(r => ({
            contacc: r.contacc || '',
            buspartner: r.buspartner || '',
            validity: r.validity || '',
            replacementCA: r.replacementCA || '',
            est12month: r.est12month || '',
            est12monthMetered: r.est12monthMetered || '',  // ← ADD
            est12monthUnMetered: r.est12monthUnMetered || '',  // ← ADD
            RECBlender_flag: r.RECBlender_flag || '',  // ← ADD
            providercontractId: r.providercontractId || '',
            cycle20: r.cycle20 || '',
            metered: r.metered || '',
            arrear60: r.arrear60 || '',
            nsfFlag: r.nsfFlag || '',
            inactive: r.inactive || '',
            enrolled: r.enrolled || '',
            selected: r.selected === 'true' || r.selected === true
          })),

          prospects: prList.map(r => ({
            siteAddLoc: r.siteAddLoc || '',
            projectedCon: r.projectedCon || '',
            year: r.year || ''
          }))
        };

        return JSON.stringify({ value: JSON.stringify(payload) });

      } catch (e) {
        req.error(500, e.message || String(e));
      }
    });

    // ── getAllOppIds ────────────────────────────────────────────────────────
    srv.on('getAllOppIds', async (req) => {
      try {
        const db = await cds.connect.to('db');
        const { Opportunity } = db.entities('com.salescloud');
        const rows = await db.run(
          SELECT.from(Opportunity)
            .columns('Oppid', 'oppCategory', 'status', 'modifiedAt')
            .where({ objectStatus: 'OPP' })
            .orderBy({ modifiedAt: 'desc' })
        );

        return JSON.stringify(
          rows.map(r => ({
            Oppid: r.Oppid,
            oppCategory: r.oppCategory || '',
            status: r.status || '',
            modifiedAt: r.modifiedAt || ''
          }))
        );
      } catch (e) {
        req.error(500, e.message || String(e));
      }
    });

    // ── getOrReplicateQuote ─────────────────────────────────────────────
    // Case 2: Check if QuoteID exists → load it, else replicate from OppID
    // ── getOrReplicateQuote ─────────────────────────────────────────────
    // Case 2: Check if QuoteID exists → load it, else replicate from OppID
    // Case 3: oppId empty → search by quoteId alone (QuoteID-only URL)
    srv.on('getOrReplicateQuote', async (req) => {
      const { oppId, quoteId } = req.data;

      // FIX 1: allow oppId to be empty (Case 3: QuoteID-only URL)
      if (!quoteId) {
        return JSON.stringify({ value: 'error', message: 'quoteId is required' });
      }

      try {
        const db = await cds.connect.to('db');
        const { Opportunity, InvolvedParties, ProviderContract,
          ConsumptionDetail, Prospect } = db.entities('com.salescloud');

        // Step 1: Check if Quote already exists
        // FIX 2: if oppId is empty, search by quoteId alone across all QUOTE rows
        const existingQuote = await db.run(
          oppId && oppId.trim()
            ? SELECT.one.from(Opportunity)
              .where({ Oppid: oppId.trim(), objectStatus: 'QUOTE', quoteId: quoteId.trim() })
            : SELECT.one.from(Opportunity)
              .where({ objectStatus: 'QUOTE', quoteId: quoteId.trim() })
        );

        if (existingQuote) {
          // Quote exists — load it directly
          console.log('Quote exists, loading:', quoteId);
          // ── ADD THIS BLOCK ────────────────────────────────────────────────
          const existingPC = await db.run(
            SELECT.from(ProviderContract).where({ opportunity_ID: existingQuote.ID })
          );

          if (oppId && oppId.trim()) {
            const opp = await db.run(
              SELECT.one.from(Opportunity)
                .where({ Oppid: oppId.trim(), objectStatus: 'OPP' })
            );

            if (opp) {
              const quoteKeys = existingPC.map(r => `${r.product}|${r.portfolio}|${r.startDate}`);

              const oppPC = await db.run(
                SELECT.from(ProviderContract).where({ opportunity_ID: opp.ID })
              );

              const allKeys = oppPC.map(r => `${r.product}|${r.portfolio}|${r.startDate}`);
              const duplicates = allKeys.filter((k, i) => allKeys.indexOf(k) !== i);
              if (duplicates.length > 0) {
                console.warn('Duplicate product+portfolio combos in OPP PC rows:', duplicates);
              }

              const newRows = oppPC.filter(r =>
                !quoteKeys.includes(`${r.product}|${r.portfolio}|${r.startDate}`)
              );

              if (newRows.length > 0) {
                const copyRows = newRows.map(r => ({
                  ...r,
                  ID: cds.utils.uuid(),
                  opportunity_ID: existingQuote.ID  // ← existingQuote.ID not quoteShell.ID
                }));
                await db.run(INSERT.into(ProviderContract).entries(copyRows));
                console.log(`Merged ${newRows.length} new PC rows into Quote:`, quoteId);
              }
            }
          }
          // ── END OF ADDED BLOCK ───────────────────────────────────────────

          const [ip, pc, cd, pr] = await Promise.all([
            db.run(SELECT.from(InvolvedParties).where({ opportunity_ID: existingQuote.ID })),
            db.run(SELECT.from(ProviderContract).where({ opportunity_ID: existingQuote.ID })),
            db.run(SELECT.from(ConsumptionDetail).where({ opportunity_ID: existingQuote.ID })),
            db.run(SELECT.from(Prospect).where({ opportunity_ID: existingQuote.ID }))
          ]);

          // Fall back to parent OPP status if Quote shell has none
          // let sQuoteStatus = existingQuote.status || '';
          // if (!sQuoteStatus && oppId && oppId.trim()) {
          //   const oppForStatus = await db.run(
          //     SELECT.one.from(Opportunity).columns('status')
          //       .where({ Oppid: oppId.trim(), objectStatus: 'OPP' })
          //   );
          //   sQuoteStatus = (oppForStatus && oppForStatus.status) ? oppForStatus.status : '';
          // }
          // Quote status is always "Open" — never inherit from OPP
          // (OPP may be Structuring/Approved etc. which is not valid for a freshly loaded Quote)
          let sQuoteStatus = existingQuote.status || 'Open';

          const result = {
            _isNew: false,
            ID: existingQuote.ID,
            Oppid: existingQuote.Oppid,
            quoteId: existingQuote.quoteId || '',
            objectStatus: 'QUOTE',
            oppCategory: existingQuote.oppCategory || '',
            // status: existingQuote.status || '',
            status: sQuoteStatus,
            salesPhase: existingQuote.salesPhase || '',
            annualGross: existingQuote.annualGross || '',
            sizeofOpp: existingQuote.sizeofOpp || '',
            percentUsage: existingQuote.percentUsage || '85',
            annualSubs: existingQuote.annualSubs || '',
            oppSubs: existingQuote.oppSubs || '',
            contLifetimeval: existingQuote.contLifetimeval || '',
            salesType: existingQuote.salesType || '',
            priceStructure: existingQuote.priceStructure || '',
            nte: existingQuote.nte || '',
            term: existingQuote.term || '',
            meteredCon: existingQuote.meteredCon || '',
            unmeteredCon: existingQuote.unmeteredCon || '',
            commencementLetterSent: existingQuote.commencementLetterSent || '',
            commencementLetterSigned: existingQuote.commencementLetterSigned || '',
            enrolled: existingQuote.enrolled || '',
            includeInPipeline: existingQuote.includeInPipeline || '',
            salesPhase: existingQuote.salesPhase || '',
            lastFetchedAt: existingQuote.lastFetchedAt || null,
            involvedParties: ip.map(r => ({ buspartner: r.buspartner || '', role: r.role || '' })),
            providerContracts: pc.map(r => ({
              product: r.product || '', portfolio: r.portfolio || '',
              fixedMWh: r.fixedMWh || '', fixedPrice: r.fixedPrice || '',
              portfolioprice: r.portfolioprice || '', subspercent: r.subspercent || '',
              startDate: r.startDate || '', endDate: r.endDate || '',
              estUsage: r.estUsage || '', recPrice: r.recPrice || '',
              migpPercent: r.migpPercent || '', recpercent: r.recpercent || '',
              migpMWh: r.migpMWh || '', recMWh: r.recMWh || '',
              enrollChk: r.enrollChk === true || r.enrollChk === 'true',
              exportDate: r.exportDate || '',
              exportChk: r.exportChk === true || r.exportChk === 'true',
              netPremium: r.netPremium || ''   // ← ADD
            })),
            consumptionDetails: cd.map(r => ({
              contacc: r.contacc || '', buspartner: r.buspartner || '',
              validity: r.validity || '', replacementCA: r.replacementCA || '',
              est12month: r.est12month || '',
              est12monthMetered: r.est12monthMetered || '',
              est12monthUnMetered: r.est12monthUnMetered || '',
              RECBlender_flag: r.RECBlender_flag || '',
              providercontractId: r.providercontractId || '',
              cycle20: r.cycle20 || '', metered: r.metered || '',
              arrear60: r.arrear60 || '', nsfFlag: r.nsfFlag || '',
              inactive: r.inactive || '', enrolled: r.enrolled || '',
              selected: r.selected === 'true' || r.selected === true
            })),
            prospects: pr.map(r => ({
              siteAddLoc: r.siteAddLoc || '',
              projectedCon: r.projectedCon || '',
              year: r.year || ''
            }))
          };

          return JSON.stringify({ value: JSON.stringify(result) });
        }

        // Step 2: Quote does NOT exist — replicate from OppID
        // FIX 3: if oppId is empty here, we cannot replicate — return helpful error
        if (!oppId || !oppId.trim()) {
          return JSON.stringify({
            value: 'error',
            message: 'Quote ' + quoteId + ' not found in BTP. Please open this quote from its linked Opportunity first so it can be replicated.'
          });
        }

        console.log('Quote not found, replicating from OppID:', oppId);

        const opp = await db.run(
          SELECT.one.from(Opportunity)
            .where({ Oppid: oppId, objectStatus: 'OPP' })
        );

        if (!opp) {
          return JSON.stringify({ value: 'error', message: 'Opportunity ' + oppId + ' not found' });
        }

        const [ip, pc, cd, pr] = await Promise.all([
          db.run(SELECT.from(InvolvedParties).where({ opportunity_ID: opp.ID })),
          db.run(SELECT.from(ProviderContract).where({ opportunity_ID: opp.ID })),
          db.run(SELECT.from(ConsumptionDetail).where({ opportunity_ID: opp.ID })),
          db.run(SELECT.from(Prospect).where({ opportunity_ID: opp.ID }))
        ]);

        // Step 3: Create new Quote record (copy of Opp, same Oppid)
        const newQuoteID = cds.utils.uuid();

        await db.run(INSERT.into(Opportunity).entries({
          ID: newQuoteID,
          Oppid: oppId,            // ← SAME as Opportunity
          objectStatus: 'QUOTE',   // ← mark as Quote
          quoteId: quoteId,        // ← stored here
          oppCategory: opp.oppCategory || '',
          status: 'Open',            // ← Quote always starts as Open, never copied from Opp
          salesPhase: opp.salesPhase || '',
          annualGross: opp.annualGross || '',
          sizeofOpp: opp.sizeofOpp || '',
          percentUsage: opp.percentUsage || '85',
          annualSubs: opp.annualSubs || '',
          oppSubs: opp.oppSubs || '',
          contLifetimeval: opp.contLifetimeval || '',
          salesType: opp.salesType || '',
          priceStructure: opp.priceStructure || '',
          nte: opp.nte || '',
          term: opp.term || '',
          meteredCon: opp.meteredCon || '',
          unmeteredCon: opp.unmeteredCon || '',
          commencementLetterSent: opp.commencementLetterSent || '',
          commencementLetterSigned: opp.commencementLetterSigned || '',
          enrolled: opp.enrolled || '',
          includeInPipeline: opp.includeInPipeline || '',
          salesPhase: opp.salesPhase || '',
          enrollmentReferralCode: opp.enrollmentReferralCode || '',  // FIX: was missing
          enrollmentEmailId: opp.enrollmentEmailId || '',            // FIX: was missing
          lastFetchedAt: opp.lastFetchedAt || null                   // FIX: was missing
        }));

        // Step 4: Copy all 4 child tables
        const copyRows = (rows) => rows.map(r => {
          const copy = Object.assign({}, r);
          // delete copy.ID;
          // delete copy.opportunity_ID;
          copy.ID = cds.utils.uuid();        // ← ALWAYS assign a fresh UUID
          copy.opportunity_ID = newQuoteID;
          copy.opportunity_ID = newQuoteID;
          delete copy.createdAt;
          delete copy.modifiedAt;
          delete copy.createdBy;
          delete copy.modifiedBy;
          return copy;
        });

        if (ip.length > 0) await db.run(INSERT.into(InvolvedParties).entries(copyRows(ip)));
        if (pc.length > 0) await db.run(INSERT.into(ProviderContract).entries(copyRows(pc)));
        if (cd.length > 0) await db.run(INSERT.into(ConsumptionDetail).entries(copyRows(cd)));
        if (pr.length > 0) await db.run(INSERT.into(Prospect).entries(copyRows(pr)));

        console.log('Replicated OppID', oppId, '→ QuoteID', quoteId);

        // Step 5: Return new Quote data
        const result = {
          _isNew: true,
          ID: newQuoteID,
          Oppid: oppId,
          objectStatus: 'QUOTE',
          quoteId: quoteId,
          oppCategory: opp.oppCategory || '',
          status: 'Open',            // ← Quote always starts as Open
          salesPhase: opp.salesPhase || '',
          annualGross: opp.annualGross || '',
          sizeofOpp: opp.sizeofOpp || '',
          percentUsage: opp.percentUsage || '85',
          annualSubs: opp.annualSubs || '',
          oppSubs: opp.oppSubs || '',
          contLifetimeval: opp.contLifetimeval || '',
          salesType: opp.salesType || '',
          priceStructure: opp.priceStructure || '',
          nte: opp.nte || '',
          term: opp.term || '',
          meteredCon: opp.meteredCon || '',
          unmeteredCon: opp.unmeteredCon || '',
          commencementLetterSent: opp.commencementLetterSent || '',
          commencementLetterSigned: opp.commencementLetterSigned || '',
          enrolled: opp.enrolled || '',
          includeInPipeline: opp.includeInPipeline || '',
          salesPhase: opp.salesPhase || '',
          lastFetchedAt: opp.lastFetchedAt || null,
          involvedParties: ip.map(r => ({ buspartner: r.buspartner || '', role: r.role || '' })),
          providerContracts: pc.map(r => ({
            product: r.product || '', portfolio: r.portfolio || '',
            fixedMWh: r.fixedMWh || '', fixedPrice: r.fixedPrice || '',
            portfolioprice: r.portfolioprice || '', subspercent: r.subspercent || '',
            startDate: r.startDate || '', endDate: r.endDate || '',
            estUsage: r.estUsage || '', recPrice: r.recPrice || '',
            migpPercent: r.migpPercent || '', recpercent: r.recpercent || '',
            migpMWh: r.migpMWh || '', recMWh: r.recMWh || '',
            enrollChk: r.enrollChk === true || r.enrollChk === 'true',
            exportDate: r.exportDate || '',
            exportChk: r.exportChk === true || r.exportChk === 'true',
            netPremium: r.netPremium || ''   // ← ADD
          })),
          consumptionDetails: cd.map(r => ({
            contacc: r.contacc || '', buspartner: r.buspartner || '',
            validity: r.validity || '', replacementCA: r.replacementCA || '',
            est12month: r.est12month || '',
            est12monthMetered: r.est12monthMetered || '',
            est12monthUnMetered: r.est12monthUnMetered || '',
            RECBlender_flag: r.RECBlender_flag || '',
            providercontractId: r.providercontractId || '',
            cycle20: r.cycle20 || '', metered: r.metered || '',
            arrear60: r.arrear60 || '', nsfFlag: r.nsfFlag || '',
            inactive: r.inactive || '', enrolled: r.enrolled || '',
            selected: r.selected === 'true' || r.selected === true
          })),
          prospects: pr.map(r => ({
            siteAddLoc: r.siteAddLoc || '',
            projectedCon: r.projectedCon || '',
            year: r.year || ''
          }))
        };

        return JSON.stringify({ value: JSON.stringify(result) });

      } catch (e) {
        console.error('getOrReplicateQuote error:', e.message);
        return JSON.stringify({ value: 'error', message: e.message });
      }
    });

    // ── getQuoteByQuoteId ────────────────────────────────────────────────────
    // Case 3: C4C sends only QuoteID in URL
    // CPI already pushed Quote shell with Oppid=456, objectStatus=QUOTE, quoteId=123
    // Step 1: Quote has data already → load directly
    // Step 2: Quote is shell only → copy data from parent Opp (same Oppid)
    // ────────────────────────────────────────────────────────────────────────
    srv.on('getQuoteByQuoteId', async (req) => {
      const { quoteId } = req.data;

      if (!quoteId || !quoteId.trim()) {
        return JSON.stringify({ value: 'error', message: 'quoteId is required' });
      }

      try {
        const db = await cds.connect.to('db');
        const { Opportunity, InvolvedParties, ProviderContract,
          ConsumptionDetail, Prospect } = db.entities('com.salescloud');

        // ── Step 1: Find Quote shell by quoteId ──────────────────────────
        // CPI pushed this with Oppid=456, objectStatus=QUOTE, quoteId=123
        const quoteShell = await db.run(
          SELECT.one.from(Opportunity)
            .where({ quoteId: quoteId.trim(), objectStatus: 'QUOTE' })
        );

        if (!quoteShell) {
          return JSON.stringify({
            value: 'error',
            message: 'Quote ' + quoteId + ' not found in BTP. ' +
              'Ensure CPI has pushed the Quote record before opening.'
          });
        }

        // ── Step 2: Check if Quote already has child data ─────────────────
        // const existingPC = await db.run(
        //   SELECT.from(ProviderContract)
        //     .where({ opportunity_ID: quoteShell.ID })
        // );
        // ── Step 2: Check if Quote already has ANY child data ─────────────
        // FIX: old code only checked ProviderContract — if OPP has no PC rows
        // this was always 0, causing repeated overwrites of quote data.
        // Now check all four tables; if any has rows the quote is populated.
        const [existingIP, existingPC, existingCD, existingPR] = await Promise.all([
          db.run(SELECT.from(InvolvedParties).where({ opportunity_ID: quoteShell.ID })),
          db.run(SELECT.from(ProviderContract).where({ opportunity_ID: quoteShell.ID })),
          db.run(SELECT.from(ConsumptionDetail).where({ opportunity_ID: quoteShell.ID })),
          db.run(SELECT.from(Prospect).where({ opportunity_ID: quoteShell.ID }))
        ]);

        if (existingIP.length > 0 || existingPC.length > 0 || existingCD.length > 0 || existingPR.length > 0) {

          // if (existingPC.length > 0) {
          // Quote already populated — load directly
          console.log('Quote already has data, loading:', quoteId);
          // ── ADD THIS BLOCK HERE ───────────────────────────────────────────
          const opp = await db.run(
            SELECT.one.from(Opportunity)
              .where({ Oppid: quoteShell.Oppid, objectStatus: 'OPP' })
          );

          if (opp) {
            const quoteKeys = existingPC.map(r => `${r.product}|${r.portfolio}|${r.startDate}`);

            const oppPC = await db.run(
              SELECT.from(ProviderContract).where({ opportunity_ID: opp.ID })
            );

            const allKeys = oppPC.map(r => `${r.product}|${r.portfolio}|${r.startDate}`);
            const duplicates = allKeys.filter((k, i) => allKeys.indexOf(k) !== i);
            if (duplicates.length > 0) {
              console.warn('Duplicate product+portfolio combos in OPP PC rows:', duplicates);
            }

            const newRows = oppPC.filter(r =>
              !quoteKeys.includes(`${r.product}|${r.portfolio}|${r.startDate}`)
            );

            if (newRows.length > 0) {
              const copyRows = newRows.map(r => ({
                ...r,
                ID: cds.utils.uuid(),
                opportunity_ID: quoteShell.ID
              }));
              await db.run(INSERT.into(ProviderContract).entries(copyRows));
              console.log(`Merged ${newRows.length} new PC rows into Quote:`, quoteId);
            }
          }
          // ── END OF ADDED BLOCK ────────────────────────────────────────────

          const [ip, pc, cd, pr] = await Promise.all([
            db.run(SELECT.from(InvolvedParties).where({ opportunity_ID: quoteShell.ID })),
            db.run(SELECT.from(ProviderContract).where({ opportunity_ID: quoteShell.ID })),
            db.run(SELECT.from(ConsumptionDetail).where({ opportunity_ID: quoteShell.ID })),
            db.run(SELECT.from(Prospect).where({ opportunity_ID: quoteShell.ID }))
          ]);


          // Fall back to parent OPP status if Quote shell has none
          // let sQuoteStatus = quoteShell.status || '';
          // if (!sQuoteStatus && quoteShell.Oppid) {
          //   const oppForStatus = await db.run(
          //     SELECT.one.from(Opportunity).columns('status')
          //       .where({ Oppid: quoteShell.Oppid, objectStatus: 'OPP' })
          //   );
          //   sQuoteStatus = (oppForStatus && oppForStatus.status) ? oppForStatus.status : '';
          // }
          // Quote status is always "Open" — never inherit from OPP
          let sQuoteStatus = quoteShell.status || 'Open';

          const result = {
            _isNew: false,
            ID: quoteShell.ID,
            Oppid: quoteShell.Oppid,
            quoteId: quoteShell.quoteId || '',
            objectStatus: 'QUOTE',
            oppCategory: quoteShell.oppCategory || '',
            status: sQuoteStatus,
            // status: quoteShell.status || '',
            salesPhase: quoteShell.salesPhase || '',
            annualGross: quoteShell.annualGross || '',
            sizeofOpp: quoteShell.sizeofOpp || '',
            percentUsage: quoteShell.percentUsage || '85',
            annualSubs: quoteShell.annualSubs || '',
            oppSubs: quoteShell.oppSubs || '',
            contLifetimeval: quoteShell.contLifetimeval || '',
            salesType: quoteShell.salesType || '',
            priceStructure: quoteShell.priceStructure || '',
            nte: quoteShell.nte || '',
            term: quoteShell.term || '',
            meteredCon: quoteShell.meteredCon || '',
            unmeteredCon: quoteShell.unmeteredCon || '',
            commencementLetterSent: quoteShell.commencementLetterSent || '',
            commencementLetterSigned: quoteShell.commencementLetterSigned || '',
            enrolled: quoteShell.enrolled || '',
            enrollmentReferralCode: quoteShell.enrollmentReferralCode || '',
            enrollmentEmailId: quoteShell.enrollmentEmailId || '',
            includeInPipeline: quoteShell.includeInPipeline || '',
            lastFetchedAt: quoteShell.lastFetchedAt || null,
            involvedParties: ip.map(r => ({
              buspartner: r.buspartner || '', role: r.role || ''
            })),
            providerContracts: pc.map(r => ({
              product: r.product || '', portfolio: r.portfolio || '',
              fixedMWh: r.fixedMWh || '', fixedPrice: r.fixedPrice || '',
              portfolioprice: r.portfolioprice || '', subspercent: r.subspercent || '',
              startDate: r.startDate || '', endDate: r.endDate || '',
              estUsage: r.estUsage || '', recPrice: r.recPrice || '',
              migpPercent: r.migpPercent || '', recpercent: r.recpercent || '',
              migpMWh: r.migpMWh || '', recMWh: r.recMWh || '',
              enrollChk: r.enrollChk === true || r.enrollChk === 'true',
              exportChk: r.exportChk === true || r.exportChk === 'true',
              exportDate: r.exportDate || '',
              netPremium: r.netPremium || ''  // ← ADD
            })),
            consumptionDetails: cd.map(r => ({
              contacc: r.contacc || '', buspartner: r.buspartner || '',
              validity: r.validity || '', replacementCA: r.replacementCA || '',
              est12month: r.est12month || '',
              est12monthMetered: r.est12monthMetered || '',
              est12monthUnMetered: r.est12monthUnMetered || '',
              RECBlender_flag: r.RECBlender_flag || '',
              providercontractId: r.providercontractId || '',
              cycle20: r.cycle20 || '', metered: r.metered || '',
              arrear60: r.arrear60 || '', nsfFlag: r.nsfFlag || '',
              inactive: r.inactive || '', enrolled: r.enrolled || '',
              selected: r.selected === 'true' || r.selected === true
            })),
            prospects: pr.map(r => ({
              siteAddLoc: r.siteAddLoc || '',
              projectedCon: r.projectedCon || '',
              year: r.year || ''
            }))
          };

          return JSON.stringify({ value: JSON.stringify(result) });
        }

        // ── Step 3: First open — copy from parent Opp (same Oppid) ───────
        // quoteShell.Oppid = 456 (same as parent Opportunity)
        console.log('First open — copying from Opp:', quoteShell.Oppid, '→ Quote:', quoteId);

        const opp = await db.run(
          SELECT.one.from(Opportunity)
            .where({ Oppid: quoteShell.Oppid, objectStatus: 'OPP' })
        );

        if (!opp) {
          return JSON.stringify({
            value: 'error',
            message: 'Parent Opportunity ' + quoteShell.Oppid + ' not found in BTP. ' +
              'Please open the Opportunity first to save its data.'
          });
        }

        // Fetch all 4 child tables from parent Opp
        const [ip, pc, cd, pr] = await Promise.all([
          db.run(SELECT.from(InvolvedParties).where({ opportunity_ID: opp.ID })),
          db.run(SELECT.from(ProviderContract).where({ opportunity_ID: opp.ID })),
          db.run(SELECT.from(ConsumptionDetail).where({ opportunity_ID: opp.ID })),
          db.run(SELECT.from(Prospect).where({ opportunity_ID: opp.ID }))
        ]);

        // Update Quote shell with full Opp header data
        await db.run(
          UPDATE(Opportunity).set({
            oppCategory: opp.oppCategory || '',
            status: 'Open',            // ← Quote always starts as Open, never copied from Opp
            salesPhase: opp.salesPhase || '',
            annualGross: opp.annualGross || '',
            sizeofOpp: opp.sizeofOpp || '',
            percentUsage: opp.percentUsage || '85',
            annualSubs: opp.annualSubs || '',
            oppSubs: opp.oppSubs || '',
            contLifetimeval: opp.contLifetimeval || '',
            salesType: opp.salesType || '',
            priceStructure: opp.priceStructure || '',
            nte: opp.nte || '',
            annualMWh: opp.annualMWh || '',   // FIX: was missing
            term: opp.term || '',
            meteredCon: opp.meteredCon || '',
            unmeteredCon: opp.unmeteredCon || '',
            commencementLetterSent: opp.commencementLetterSent || '',
            commencementLetterSigned: opp.commencementLetterSigned || '',
            enrolled: opp.enrolled || '',
            enrollmentReferralCode: opp.enrollmentReferralCode || '',
            enrollmentEmailId: opp.enrollmentEmailId || '',
            includeInPipeline: opp.includeInPipeline || '',
            lastFetchedAt: opp.lastFetchedAt
          }).where({ ID: quoteShell.ID })
        );

        // Copy child tables from Opp → Quote
        const copyRows = (rows) => rows.map(r => {
          const copy = Object.assign({}, r);
          // delete copy.ID;
          // delete copy.opportunity_ID;
          copy.ID = cds.utils.uuid();        // ← ALWAYS assign a fresh UUID
          copy.opportunity_ID = quoteShell.ID;
          copy.opportunity_ID = quoteShell.ID;
          delete copy.createdAt;
          delete copy.modifiedAt;
          delete copy.createdBy;
          delete copy.modifiedBy;
          return copy;
        });

        if (ip.length > 0) await db.run(INSERT.into(InvolvedParties).entries(copyRows(ip)));
        if (pc.length > 0) await db.run(INSERT.into(ProviderContract).entries(copyRows(pc)));
        if (cd.length > 0) await db.run(INSERT.into(ConsumptionDetail).entries(copyRows(cd)));
        if (pr.length > 0) await db.run(INSERT.into(Prospect).entries(copyRows(pr)));

        console.log('Quote', quoteId, 'populated from Opp', quoteShell.Oppid);

        // Return populated Quote
        const result = {
          _isNew: true,
          ID: quoteShell.ID,
          Oppid: quoteShell.Oppid,
          quoteId: quoteId,
          objectStatus: 'QUOTE',
          oppCategory: opp.oppCategory || '',
          status: 'Open',            // ← Quote always starts as Open
          salesPhase: opp.salesPhase || '',
          annualGross: opp.annualGross || '',
          sizeofOpp: opp.sizeofOpp || '',
          percentUsage: opp.percentUsage || '85',
          annualSubs: opp.annualSubs || '',
          oppSubs: opp.oppSubs || '',
          contLifetimeval: opp.contLifetimeval || '',
          salesType: opp.salesType || '',
          priceStructure: opp.priceStructure || '',
          nte: opp.nte || '',
          term: opp.term || '',
          meteredCon: opp.meteredCon || '',
          unmeteredCon: opp.unmeteredCon || '',
          commencementLetterSent: opp.commencementLetterSent || '',
          commencementLetterSigned: opp.commencementLetterSigned || '',
          enrolled: opp.enrolled || '',
          enrollmentReferralCode: opp.enrollmentReferralCode || '',
          enrollmentEmailId: opp.enrollmentEmailId || '',
          includeInPipeline: opp.includeInPipeline || '',
          lastFetchedAt: opp.lastFetchedAt || null,
          involvedParties: ip.map(r => ({
            buspartner: r.buspartner || '', role: r.role || ''
          })),
          providerContracts: pc.map(r => ({
            product: r.product || '', portfolio: r.portfolio || '',
            fixedMWh: r.fixedMWh || '', fixedPrice: r.fixedPrice || '',
            portfolioprice: r.portfolioprice || '', subspercent: r.subspercent || '',
            startDate: r.startDate || '', endDate: r.endDate || '',
            estUsage: r.estUsage || '', recPrice: r.recPrice || '',
            migpPercent: r.migpPercent || '', recpercent: r.recpercent || '',
            migpMWh: r.migpMWh || '', recMWh: r.recMWh || '',
            enrollChk: r.enrollChk === true || r.enrollChk === 'true',
            exportChk: r.exportChk === true || r.exportChk === 'true',
            exportDate: r.exportDate || '',
            netPremium: r.netPremium || ''   // ← ADD
          })),
          consumptionDetails: cd.map(r => ({
            contacc: r.contacc || '', buspartner: r.buspartner || '',
            validity: r.validity || '', replacementCA: r.replacementCA || '',
            est12month: r.est12month || '',
            est12monthMetered: r.est12monthMetered || '',
            est12monthUnMetered: r.est12monthUnMetered || '',
            RECBlender_flag: r.RECBlender_flag || '',
            providercontractId: r.providercontractId || '',
            cycle20: r.cycle20 || '', metered: r.metered || '',
            arrear60: r.arrear60 || '', nsfFlag: r.nsfFlag || '',
            inactive: r.inactive || '', enrolled: r.enrolled || '',
            selected: r.selected === 'true' || r.selected === true
          })),
          prospects: pr.map(r => ({
            siteAddLoc: r.siteAddLoc || '',
            projectedCon: r.projectedCon || '',
            year: r.year || ''
          }))
        };

        return JSON.stringify({ value: JSON.stringify(result) });

      } catch (e) {
        console.error('getQuoteByQuoteId error:', e.message);
        return JSON.stringify({ value: 'error', message: e.message });
      }
    });

    //-fetchValidateCAs_____________________________
    srv.on('fetchValidateCAs', async (req) => {
      // ── ADD THIS ─────────────────────────────────────────
      console.log('[fetchValidateCAs] Raw req.data:', JSON.stringify(req.data));
      console.log('[fetchValidateCAs] businessPartners raw:', req.data.businessPartners);
      console.log('[fetchValidateCAs] selectedCAs raw:', req.data.selectedCAs);
      // ─────────────────────────────────────────────────────
      const { businessPartners, selectedCAs, oppId, objectStatus, quoteId } = req.data;

      try {
        const aBPs = JSON.parse(businessPartners || '[]');
        const aSelectedCAs = JSON.parse(selectedCAs || '[]'); // array of contacc strings
        const bIsFirstFetch = (aSelectedCAs.length === 0);

        // Step 1: Get OAuth Token
        // 'Salescloud_CRB_CPI' must match the destination name in BTP Cockpit exactly
        // const cpiDest = await cds.connect.to('Salescloud_CRB_CPI');
        // ── Step 1: Resolve destination ──────────────────────────────────
        let cpiDest;
        try {
          cpiDest = await cds.connect.to('Salescloud_CRB_CPI');
          console.log('[fetchValidateCAs] ✅ Destination resolved: Salescloud_CRB_CPI');
        } catch (destErr) {
          console.error('[fetchValidateCAs] ❌ Destination resolution failed:', destErr.message);
          // This message surfaces to controller → shown as MessageBox
          req.error(502, 'DEST_NOT_FOUND: Could not resolve destination Salescloud_CRB_CPI. ' + destErr.message);
          return;
        }
        // Step 2: Build CPI payload
        // First Fetch  → send BPs only, empty contractAccounts → CPI returns all active CAs
        // Next Fetches → send BPs + selected CAs → CPI checks inactivity + returns all active CAs
        const cpiPayload = {
          businessPartners: aBPs.map(bp => ({
            businessPartnerId: bp.businessPartnerId,
            contractAccounts: bIsFirstFetch ? [] : aSelectedCAs
          }))
        };

        console.log('[fetchValidateCAs] Sending to CPI → payload:', JSON.stringify(cpiPayload));
        // // Step 3: Send via destination — BTP injects Basic Auth header automatically
        // const data = await cpiDest.send({
        //   method: 'POST',
        //   path: '/',           // adjust if your CPI iFlow has a specific path suffix
        //   data: cpiPayload,
        //   headers: { 'Content-Type': 'application/json' }
        // });
        // console.log('[fetchValidateCAs] CPI raw response:', JSON.stringify(data));
        // ── Step 3: Call CPI ─────────────────────────────────────────────
        let data;
        try {
          data = await cpiDest.send({
            method: 'POST',
            path: '/',
            data: cpiPayload,
            headers: { 'Content-Type': 'application/json' }
          });
          console.log('[fetchValidateCAs] ✅ CPI responded:', JSON.stringify(data));
        } catch (cpiErr) {
          console.error('[fetchValidateCAs] ❌ CPI call failed:', cpiErr.message);
          // Distinguish between auth failure vs URL not reachable
          const status = cpiErr.response?.status;
          if (status === 401 || status === 403) {
            req.error(502, 'CPI_AUTH_FAILED: Destination credentials rejected by CPI. HTTP ' + status);
          } else if (status === 404) {
            req.error(502, 'CPI_NOT_FOUND: CPI iFlow URL not found. HTTP 404 — check destination URL.');
          } else if (!status) {
            req.error(502, 'CPI_UNREACHABLE: No response from CPI — check destination URL or network.');
          } else {
            req.error(502, 'CPI_ERROR: CPI returned HTTP ' + status + '. ' + cpiErr.message);
          }
          return;
        }
        // CPI might return contractAccounts or businessPartners — handle both
        var aFreshCAs = data.contractAccounts || [];

        console.log('[fetchValidateCAs] contractAccounts from CPI:', JSON.stringify(aFreshCAs));
        console.log('[fetchValidateCAs] Full CPI response keys:', Object.keys(data));
        console.log('[fetchValidateCAs] CA count from CPI:', aFreshCAs.length);

        // Step 4: Persist lastFetchedAt to correct OPP or QUOTE row
        let lastFetchedAt = null;
        if (oppId && oppId.trim()) {
          try {
            const db = await cds.connect.to('db');
            const { Opportunity } = db.entities('com.salescloud');
            lastFetchedAt = new Date().toISOString();

            const isQuote = (objectStatus === 'QUOTE');
            const whereClause = isQuote
              ? { Oppid: oppId.trim(), objectStatus: 'QUOTE', quoteId }
              : { Oppid: oppId.trim(), objectStatus: 'OPP' };

            await db.run(UPDATE(Opportunity).set({ lastFetchedAt }).where(whereClause));
            console.log('lastFetchedAt saved for',
              isQuote ? 'QuoteId:' + quoteId : 'OppId:' + oppId, '→', lastFetchedAt);
          } catch (dbErr) {
            console.warn('Failed to persist lastFetchedAt:', dbErr.message);
          }
        }

        // Step 5: Return CA data + metadata to controller
        return JSON.stringify({
          contractAccounts: data.contractAccounts || [],
          isFirstFetch: bIsFirstFetch,
          lastFetchedAt: lastFetchedAt
        });

      } catch (e) {
        console.error('[fetchValidateCAs] ❌ Unexpected error:', e.message);
        req.error(500, 'CR&B fetch failed: ' + e.message);
      }
    });

    // ── pushSalesPhaseToC4C ─────────────────────────────────────────────────
    // Flow 2: BTP → CPI → C4C
    // Triggered when user changes Sales Phase in BTP after Structuring
    // ⚠️ Replace placeholder URLs with actual values from CPI team
    // ───────────────────────────────────────────────────────────────────────
    srv.on('pushSalesPhaseToC4C', async (req) => {
      try {
        const payload = JSON.parse(req.data.bundle || '{}');
        const {
          oppId,
          quoteId,
          objectStatus,
          salesPhase,
          salesPhaseText,
          status,
          oppCategory
        } = payload;

        if (!oppId || !salesPhase) {
          return JSON.stringify({
            value: 'error',
            message: 'oppId and salesPhase are required'
          });
        }

        const sPhaseText = oPhaseMap[salesPhase] || salesPhaseText || '';
        console.log('pushSalesPhaseToC4C called:', {
          oppId,
          salesPhase,
          salesPhaseText: sPhaseText,
          status,
          oppCategory
        });

        // ── BTP Destination → CPI iFlow → C4C ────────────────────────
        const dest = await cds.connect.to('salescloud_phase_cpi');
        // In quote context send quoteId so C4C updates the Quote, not the Opportunity
        const isQuote = (objectStatus === 'QUOTE');
        const cpiPayload = {
          // opportunityId: isQuote ? quoteId : oppId,  // ← quote uses quoteId
          opportunityId: oppId,
          ...(isQuote && quoteId ? { quoteId: quoteId } : {}),
          salesPhaseCode: salesPhase,   // e.g. "Z0"
          salesPhaseText: sPhaseText,   // e.g. "Approval"
          status: status,
          oppCategory: oppCategory
        };

        console.log('pushSalesPhaseToC4C — sending via destination:', cpiPayload);

        const response = await dest.send({
          method: 'POST',
          path: '/',
          data: cpiPayload
        });

        console.log('CPI iFlow response:', response ?? '(empty body - success)');

        return JSON.stringify({
          value: 'ok',
          message: 'Sales Phase pushed to C4C successfully'
        });

        // } catch (e) {
        //   console.error('pushSalesPhaseToC4C error:', e.message);
        //   return JSON.stringify({
        //     value: 'warning',
        //     message: 'BTP updated. C4C sync failed: ' + e.message
        //   });
        // }
      } catch (e) {
        // ── Extract the real CPI error body from the 500 response ──────────
        let sCPIError = e.message || 'Unknown error';
        try {
          // CAP wraps remote errors — body is in e.response?.data or e.cause?.response?.data
          const oErrData = e.response?.data ?? e.cause?.response?.data ?? e.cause ?? null;
          if (oErrData) {
            const sBody = typeof oErrData === 'string' ? oErrData : JSON.stringify(oErrData);
            sCPIError = sBody;
            console.error('pushSalesPhaseToC4C — CPI error body:', sBody);
          }
        } catch (_) { /* ignore secondary parse failure */ }

        console.error('pushSalesPhaseToC4C message:', e.message);
        console.error('pushSalesPhaseToC4C stack:', e.stack);

        return JSON.stringify({
          value: 'warning',
          message: 'BTP updated. C4C sync failed: ' + sCPIError
        });
      }
    });

    // ── initiateEnrollment ──────────────────────────────────────────────────
    // BTP → CPI → CR&B
    // Called from Enroll button in Provider Contract tab
    // Requires: enrollmentReferralCode + enrollmentEmailId at header
    //           enrollChk checked on PC rows
    //           CA rows selected
    // ───────────────────────────────────────────────────────────────────────
    srv.on('initiateEnrollment', async (req) => {
      try {
        const payload = JSON.parse(req.data.bundle || '{}');
        const {
          oppId,
          quoteId,
          objectStatus,
          referralCode,
          emailId,
          providerContracts,
          contractAccounts
        } = payload;

        // ── Validate required fields ──────────────────────────────────────
        if (!referralCode || !referralCode.trim()) {
          return JSON.stringify({
            value: 'error',
            message: 'Enrollment Referral Code is required'
          });
        }
        if (!emailId || !emailId.trim()) {
          return JSON.stringify({
            value: 'error',
            message: 'Enrollment Email ID is required'
          });
        }
        if (!providerContracts || providerContracts.length === 0) {
          return JSON.stringify({
            value: 'error',
            message: 'No Provider Contract rows selected for enrollment'
          });
        }
        if (!contractAccounts || contractAccounts.length === 0) {
          return JSON.stringify({
            value: 'error',
            message: 'No Contract Accounts selected for enrollment'
          });
        }

        console.log('initiateEnrollment called:', {
          oppId,
          quoteId,
          objectStatus,
          referralCode,
          emailId,
          pcCount: providerContracts.length,
          caCount: contractAccounts.length
        });

        // ── TEMPORARY MOCK — remove when CPI iFlow is ready ────────────
        console.log('initiateEnrollment — waiting for CPI enrollment iFlow');
        return JSON.stringify({
          value: 'ok',
          message: 'Enrollment initiated (mock)',
          providerContractId: ''  // CR&B will return actual contract ID
        });

        // ── BTP Destination block — uncomment when ready ────────────────
        /*
        const dest = await cds.connect.to('salescloud_crb_cpi');
        const response = await dest.send({
          method : 'POST',
          path   : '/enrollment',  // confirm path with CPI team
          data   : {
            opportunityId   : oppId,
            referralCode    : referralCode,
            emailId         : emailId,
            contractAccounts: contractAccounts.map(ca => ({
              contractAccountId: ca.contacc,
              businessPartnerId: ca.buspartner
            })),
            providerContracts: providerContracts.map(pc => ({
              portfolio  : pc.portfolio,
              subspercent: pc.subspercent,
              startDate  : pc.startDate,
              endDate    : pc.endDate
            }))
          }
        });
        console.log('Enrollment response:', response);
        return JSON.stringify({
          value            : 'ok',
          message          : 'Enrollment initiated successfully',
          providerContractId: response.providerContractId || ''
        });
        */

      } catch (e) {
        console.error('initiateEnrollment error:', e.message);
        return JSON.stringify({
          value: 'error',
          message: 'Enrollment failed: ' + e.message
        });
      }
    });

    // ── changeBillingCycle ──────────────────────────────────────────────────
    // BTP → CPI → CR&B
    // Called from Billing Cycle 20 button in CA Consumption Details tab
    // Changes billing cycle to 20 for selected Contract Accounts
    // ───────────────────────────────────────────────────────────────────────
    srv.on('changeBillingCycle', async (req) => {
      try {
        const payload = JSON.parse(req.data.bundle || '{}');
        const { contractAccounts } = payload;

        // ── Validate ──────────────────────────────────────────────────────
        if (!contractAccounts || contractAccounts.length === 0) {
          return JSON.stringify({
            value: 'error',
            message: 'No Contract Accounts provided for billing cycle change'
          });
        }

        console.log('changeBillingCycle called for CAs:',
          contractAccounts.map(ca => ca.contractAccountId)
        );

        // ── Update cycle20 in DB for selected CAs ─────────────────────────
        const db = await cds.connect.to('db');
        const { ConsumptionDetail } = cds.entities('com.salescloud');

        for (const ca of contractAccounts) {
          if (ca.contacc) {
            await db.run(
              UPDATE(ConsumptionDetail)
                .set({ cycle20: 'Yes' })
                .where({ contacc: ca.contacc })
            );
          }
        }

        console.log('cycle20 updated in DB for',
          contractAccounts.length, 'CAs');

        // ── TEMPORARY MOCK — remove when CPI iFlow is ready ────────────
        console.log('changeBillingCycle — waiting for CPI billing iFlow');
        return JSON.stringify({
          value: 'ok',
          message: 'Billing cycle updated in BTP. ' +
            'CR&B sync pending — waiting for CPI iFlow.'
        });

        // ── BTP Destination block — uncomment when ready ────────────────
        /*
        const dest = await cds.connect.to('salescloud_crb_cpi');
        const response = await dest.send({
          method : 'POST',
          path   : '/billingcycle',  // confirm path with CPI team
          data   : {
            contractAccounts: contractAccounts.map(ca => ({
              contractAccountId: ca.contractAccountId,
              businessPartnerId: ca.businessPartnerId,
              billingCycle     : '20'
            }))
          }
        });
        console.log('Billing cycle response:', response);
        return JSON.stringify({
          value  : 'ok',
          message: 'Billing cycle changed to 20 in CR&B'
        });
        */

      } catch (e) {
        console.error('changeBillingCycle error:', e.message);
        return JSON.stringify({
          value: 'error',
          message: 'Billing cycle change failed: ' + e.message
        });
      }
    });

    srv.on('updateLastFetchedAt', async (req) => {
      const { oppId, objectStatus, quoteId } = req.data;  // ← receive new params
      if (!oppId || !oppId.trim()) {
        return JSON.stringify({ value: 'error', message: 'oppId is required' });
      }
      try {
        const db = await cds.connect.to('db');
        const { Opportunity } = db.entities('com.salescloud');
        const lastFetchedAt = new Date().toISOString();

        const isQuote = (objectStatus === 'QUOTE');

        // Build the correct WHERE clause depending on context
        const whereClause = isQuote
          ? { Oppid: oppId.trim(), objectStatus: 'QUOTE', quoteId: quoteId }
          : { Oppid: oppId.trim(), objectStatus: 'OPP' };

        await db.run(
          UPDATE(Opportunity)
            .set({ lastFetchedAt })
            .where(whereClause)
        );
        console.log('lastFetchedAt saved for', isQuote ? 'QuoteId:' + quoteId : 'OppId:' + oppId, '→', lastFetchedAt);
        return JSON.stringify({ value: 'ok', lastFetchedAt });
      } catch (e) {
        console.error('updateLastFetchedAt error:', e.message);
        return JSON.stringify({ value: 'error', message: e.message });
      }
    });

    return;
  }

  // ─── CPIService ────────────────────────────────────────────────────────────
  if (srv.name === 'CPIService') {

    srv.on('PushOpportunityFromCPI', async (req) => {
      const { Opportunity, InvolvedParties, ProviderContract } = cds.entities('com.salescloud');
      const { opp, involvedParties, providerContracts } = req.data || {};

      if (!opp || typeof opp.Oppid !== 'string' || !opp.Oppid.trim()) {
        return req.error(400, 'opp.Oppid is required and must be a non-empty string');
      }
      if (Array.isArray(providerContracts)) {
        for (const [i, pc] of providerContracts.entries()) {
          if (pc?.startDate && pc?.endDate) {
            const sd = new Date(pc.startDate);
            const ed = new Date(pc.endDate);
            if (!isFinite(sd) || !isFinite(ed)) {
              return req.error(400, `providerContracts[${i}] has invalid dates`);
            }
            if (ed < sd) {
              return req.error(400, `providerContracts[${i}].endDate must be >= startDate`);
            }
          }
        }
      }

      // // In PushOpportunityFromCPI — add after existing validation:
      // const aValidStatuses = [
      //   "In Progress", "Structuring", "In Approval",
      //   "Approved", "Rejected", "Contract Sign-Off", "Closed", "Won"
      // ];

      // if (opp.status) {
      //   const sMatchedStatus = aValidStatuses.find(s =>
      //     s.toLowerCase() === opp.status.toLowerCase()
      //   );
      //   if (!sMatchedStatus) {
      //     console.warn('Unknown status from CPI:', opp.status);
      //     // Don't reject — just warn and save as-is
      //     // C4C may send new statuses in future
      //   } else {
      //     opp.status = sMatchedStatus; // normalize case
      //   }
      // }

      // ── Status code → label map (CPI now sends Z-codes) ──────────────────
      const oStatusCodeMap = {
        'Z0': 'Open',
        'Z1': 'In Progress',
        'Z2': 'Won',
        'Z3': 'Lost',
        'Z4': 'Closed',
        'Z5': 'Structuring'
      };

      if (opp.status) {
        // First try: direct Z-code lookup  (e.g. "Z0" → "Open")
        if (oStatusCodeMap[opp.status]) {
          opp.status = oStatusCodeMap[opp.status];
        } else {
          // Fallback: CPI still sending text — normalize case
          const aValidStatuses = [
            'Open', 'In Progress', 'Won', 'Lost', 'Closed', 'Structuring'
          ];
          const sMatched = aValidStatuses.find(s =>
            s.toLowerCase() === opp.status.toLowerCase()
          );
          if (sMatched) {
            opp.status = sMatched;
          } else {
            console.warn('Unknown status from CPI:', opp.status);
            // Don't reject — save as-is so nothing is lost
          }
        }
      }

      // ── Log received salesPhase ───────────────────────────────────────
      if (opp.salesPhase) {
        console.log('Received salesPhase from C4C:',
          opp.salesPhase, '→',
          oPhaseMap[opp.salesPhase] || 'Unknown'  // ✅ oPhaseMap used here
        );
      }

      const tx = cds.tx(req);
      let inserted = 0, updated = 0, childrenWritten = 0;

      const insertInBatches = async (Entity, rows) => {
        if (!Array.isArray(rows) || rows.length === 0) return 0;
        let ok = 0;
        for (let i = 0; i < rows.length; i += BATCH_SIZE) {
          const chunk = rows.slice(i, i + BATCH_SIZE);
          try {
            await tx.run(INSERT.into(Entity).entries(chunk));
            ok += chunk.length;
          } catch (e) {
            for (const r of chunk) {
              try {
                await tx.run(INSERT.into(Entity).entries(r));
                ok++;
              } catch (err) {
                req.warn(`Insert failed in ${Entity.name}: ${err.message || err}`);
              }
            }
          }
        }
        return ok;
      };

      try {
        // ── Route based on objectStatus ───────────────────────────────────
        // OPP  → upsert Opportunity row, replace children
        // QUOTE → create Quote shell row only (no children from CPI)
        //         children will be copied from OPP when user opens Quote iFrame

        // const isQuotePush = (opp.objectStatus === 'QUOTE');
        // ── Detect push type by quoteId presence ─────────────────────────
        // If quoteId is present → Quote shell push
        // If quoteId is absent  → Opportunity push
        // No need for CPI to send objectStatus — we derive it here
        const isQuotePush = (opp.quoteId && opp.quoteId.trim() !== '');
        // Set objectStatus accordingly — do not rely on CPI sending it
        opp.objectStatus = isQuotePush ? 'QUOTE' : 'OPP';

        console.log('PushOpportunityFromCPI — detected as:',
          opp.objectStatus,
          '| Oppid:', opp.Oppid,
          '| quoteId:', opp.quoteId || '(none)'
        );

        if (isQuotePush) {
          // ── QUOTE push — create shell row if not exists ─────────────────
          // Do NOT touch the OPP row
          // quoteId = 123, Oppid = 456 (parent link)
          console.log('Quote push received: QuoteID=', opp.quoteId,
            '← OppID=', opp.Oppid);

          const existingQuote = await tx.run(
            SELECT.one.from(Opportunity).columns('ID', 'Oppid')
              .where({ quoteId: opp.quoteId, objectStatus: 'QUOTE' })
          );

          if (existingQuote) {
            // ── GUARD: Ensure Quote belongs to the same Oppid ──────────────
            // Fetch the full row to check its parent Oppid
            const existingQuoteFull = await tx.run(
              SELECT.one.from(Opportunity).columns('ID', 'Oppid')
                .where({ quoteId: opp.quoteId, objectStatus: 'QUOTE' })
            );

            if (existingQuoteFull.Oppid !== opp.Oppid) {
              // Quote-2002 belongs to OPP-1002, but payload says OPP-1001 → REJECT
              return req.error(409,
                `Quote '${opp.quoteId}' is already linked to Opportunity '${existingQuoteFull.Oppid}'. ` +
                `Cannot re-link to '${opp.Oppid}'.`
              );
            }

            // Same Oppid — safe to update status
            await tx.run(
              UPDATE(Opportunity).set({
                status: opp.status ?? null,
                salesPhase: opp.salesPhase ?? null
              }).where({ ID: existingQuote.ID })
            );
            updated++;
            console.log('Quote shell updated:', opp.quoteId);
            // ── Merge new IP rows from payload into QUOTE ────────────────────
            // if (Array.isArray(involvedParties) && involvedParties.length > 0) {

            //   const existingIP = await tx.run(
            //     SELECT.from(InvolvedParties)
            //       .where({ opportunity_ID: existingQuote.ID })
            //   );

            //   const quoteIPKeys = existingIP.map(r => `${r.buspartner}|${r.role}`);

            //   const newIPRows = involvedParties.filter(x =>
            //     !quoteIPKeys.includes(`${x.buspartner}|${x.role}`)
            //   );

            //   if (newIPRows.length > 0) {
            //     const ipRows = newIPRows.map(x => ({
            //       ID: cds.utils.uuid(),
            //       buspartner: x?.buspartner ?? null,
            //       role: x?.role ?? null,
            //       opportunity_ID: existingQuote.ID
            //     }));
            //     await insertInBatches(InvolvedParties, ipRows);
            //     console.log(`Merged ${newIPRows.length} new IP rows into Quote:`, opp.quoteId);
            //   }
            // }
            // // ── Merge new PC rows from payload into QUOTE ────────────────────
            // if (Array.isArray(providerContracts) && providerContracts.length > 0) {

            //   const existingPC = await tx.run(
            //     SELECT.from(ProviderContract)
            //       .where({ opportunity_ID: existingQuote.ID })
            //   );

            //   const quoteKeys = existingPC.map(r => `${r.product}|${r.portfolio}`);

            //   const newRows = providerContracts.filter(x =>
            //     !quoteKeys.includes(`${x.product}|${x.portfolio}`)
            //   );

            //   if (newRows.length > 0) {
            //     const pcRows = newRows.map(x => ({
            //       ID: cds.utils.uuid(),
            //       product: x?.product ?? null,
            //       portfolio: x?.portfolio ?? null,
            //       fixedMWh: x?.fixedMWh ?? null,
            //       fixedPrice: x?.fixedPrice ?? null,
            //       portfolioprice: x?.portfolioprice ?? null,
            //       subspercent: x?.subspercent ?? null,
            //       startDate: x?.startDate ?? null,
            //       endDate: x?.endDate ?? null,
            //       enrollChk: x?.enrollChk ?? null,
            //       exportDate: x?.exportDate ?? null,
            //       exportChk: x?.exportChk ?? null,
            //       opportunity_ID: existingQuote.ID
            //     }));

            //     await insertInBatches(ProviderContract, pcRows);
            //     console.log(`Merged ${newRows.length} new PC rows into Quote:`, opp.quoteId);
            //   }
            // }
            // ── Full replace IP rows for QUOTE (C4C is source of truth) ─────
            // await tx.run(DELETE.from(InvolvedParties).where({ opportunity_ID: existingQuote.ID }));
            // if (Array.isArray(involvedParties) && involvedParties.length > 0) {
            //   const ipRows = involvedParties.map(x => ({
            //     ID: cds.utils.uuid(),
            //     buspartner: x?.buspartner ?? null,
            //     role: x?.role ?? null,
            //     opportunity_ID: existingQuote.ID
            //   }));
            //   await insertInBatches(InvolvedParties, ipRows);
            //   console.log(`Replaced ${ipRows.length} IP rows for Quote:`, opp.quoteId);
            // }

            // // ── Full replace PC rows for QUOTE (C4C is source of truth) ─────
            // await tx.run(DELETE.from(ProviderContract).where({ opportunity_ID: existingQuote.ID }));
            // if (Array.isArray(providerContracts) && providerContracts.length > 0) {
            //   const pcRows = providerContracts.map(x => ({
            //     ID: cds.utils.uuid(),
            //     product: x?.product ?? null,
            //     portfolio: x?.portfolio ?? null,
            //     fixedMWh: x?.fixedMWh ?? null,
            //     fixedPrice: x?.fixedPrice ?? null,
            //     portfolioprice: x?.portfolioprice ?? null,
            //     subspercent: x?.subspercent ?? null,
            //     startDate: x?.startDate ?? null,
            //     endDate: x?.endDate ?? null,
            //     enrollChk: x?.enrollChk ?? null,
            //     exportDate: x?.exportDate ?? null,
            //     exportChk: x?.exportChk ?? null,
            //     opportunity_ID: existingQuote.ID
            //   }));
            //   await insertInBatches(ProviderContract, pcRows);
            //   console.log(`Replaced ${pcRows.length} PC rows for Quote:`, opp.quoteId);
            // }

            // ── Replicate OPP children into QUOTE, then append payload rows ──
            // Step A: Load parent OPP
            const parentOpp = await tx.run(
              SELECT.one.from(Opportunity).columns('ID')
                .where({ Oppid: existingQuote.Oppid, objectStatus: 'OPP' })
            );

            // Step B: Wipe existing quote children — full replace
            await tx.run(DELETE.from(InvolvedParties).where({ opportunity_ID: existingQuote.ID }));
            // await tx.run(DELETE.from(ProviderContract).where({ opportunity_ID: existingQuote.ID }));
            // Snapshot QUOTE's existing exportChk/exportDate before wiping
            const existingQuotePCSnap = await tx.run(
              SELECT.from(ProviderContract).where({ opportunity_ID: existingQuote.ID })
            );
            const quoteSnapMap = {};
            existingQuotePCSnap.forEach(r => { quoteSnapMap[`${r.product}|${r.portfolio}|${r.startDate}`] = r; });

            await tx.run(DELETE.from(ProviderContract).where({ opportunity_ID: existingQuote.ID }));

            // Step C: Copy ALL IP rows from OPP into quote
            let baseIPRows = [];
            if (parentOpp) {
              const oppIP = await tx.run(
                SELECT.from(InvolvedParties).where({ opportunity_ID: parentOpp.ID })
              );
              baseIPRows = oppIP.map(r => ({
                ID: cds.utils.uuid(),
                buspartner: r.buspartner ?? null,
                role: r.role ?? null,
                opportunity_ID: existingQuote.ID
              }));
            }

            // Step D: Append NEW IP rows from payload (skip duplicates already in OPP)
            const existingIPKeys = baseIPRows.map(r => `${r.buspartner}|${r.role}`);
            const newIPRows = (involvedParties || [])
              .filter(x => !existingIPKeys.includes(`${x?.buspartner}|${x?.role}`))
              .map(x => ({
                ID: cds.utils.uuid(),
                buspartner: x?.buspartner ?? null,
                role: x?.role ?? null,
                opportunity_ID: existingQuote.ID
              }));

            const allIPRows = [...baseIPRows, ...newIPRows];
            if (allIPRows.length > 0) {
              await insertInBatches(InvolvedParties, allIPRows);
              console.log(`Quote IP: ${baseIPRows.length} from OPP + ${newIPRows.length} new from payload = ${allIPRows.length} total`);
            }

            // Step E: Copy ALL PC rows from OPP into quote
            // Step E: Copy ALL PC rows from OPP into quote
            let basePCRows = [];
            if (parentOpp) {
              const oppPC = await tx.run(
                SELECT.from(ProviderContract).where({ opportunity_ID: parentOpp.ID })
              );
              basePCRows = oppPC.map(r => {
                const snap = quoteSnapMap[`${r.product}|${r.portfolio}|${r.startDate}`] || {};
                return {
                  ID: cds.utils.uuid(),
                  product: r.product ?? null,
                  portfolio: r.portfolio ?? null,
                  startDate: r.startDate ?? null,
                  endDate: r.endDate ?? null,
                  fixedMWh: snap.fixedMWh ?? r.fixedMWh ?? null,
                  fixedPrice: snap.fixedPrice ?? r.fixedPrice ?? null,
                  portfolioprice: snap.portfolioprice ?? r.portfolioprice ?? null,
                  subspercent: snap.subspercent ?? r.subspercent ?? null,
                  enrollChk: snap.enrollChk ?? r.enrollChk ?? null,
                  exportChk: snap.exportChk ?? null,
                  exportDate: snap.exportDate ?? null,
                  // ── These were missing ──
                  recPrice: snap.recPrice ?? r.recPrice ?? null,
                  migpPercent: snap.migpPercent ?? r.migpPercent ?? null,
                  recpercent: snap.recpercent ?? r.recpercent ?? null,
                  migpMWh: snap.migpMWh ?? r.migpMWh ?? null,
                  recMWh: snap.recMWh ?? r.recMWh ?? null,
                  estUsage: snap.estUsage ?? r.estUsage ?? null,
                  netPremium: snap.netPremium ?? r.netPremium ?? null,   // ← ADD
                  opportunity_ID: existingQuote.ID
                };
              });
            }

            // Step F: Append NEW PC rows from payload (skip duplicates already in OPP)
            const existingPCKeys = basePCRows.map(r => `${r.product}|${r.portfolio}|${r.startDate}`);
            const newPCRows = (providerContracts || [])
              .filter(x => !existingPCKeys.includes(`${x?.product}|${x?.portfolio}${x?.startDate}`))
              .map(x => ({
                ID: cds.utils.uuid(),
                product: x?.product ?? null,
                portfolio: x?.portfolio ?? null,
                fixedMWh: x?.fixedMWh ?? null,
                fixedPrice: x?.fixedPrice ?? null,
                portfolioprice: x?.portfolioprice ?? null,
                subspercent: x?.subspercent ?? null,
                startDate: x?.startDate ?? null,
                endDate: x?.endDate ?? null,
                enrollChk: x?.enrollChk ?? null,
                exportDate: null,   // new rows have no prior BTP values
                exportChk: null,    // ← restore QUOTE's own value
                netPremium: x?.netPremium ?? null, // my add
                opportunity_ID: existingQuote.ID
              }));

            const allPCRows = [...basePCRows, ...newPCRows];
            if (allPCRows.length > 0) {
              await insertInBatches(ProviderContract, allPCRows);
              console.log(`Quote PC: ${basePCRows.length} from OPP + ${newPCRows.length} new from payload = ${allPCRows.length} total`);
            }
          }

          // else {
          //   // First time — create Quote shell row
          //   const newQuoteID = cds.utils.uuid();
          //   await tx.run(INSERT.into(Opportunity).entries({
          //     ID: newQuoteID,
          //     Oppid: opp.Oppid,        // ← 456 (parent OppID, used for copy)
          //     objectStatus: 'QUOTE',
          //     quoteId: opp.quoteId,      // ← 123 (C4C Quote ID)
          //     status: opp.status ?? null,
          //     oppCategory: opp.oppCategory ?? null,
          //     salesPhase: opp.salesPhase ?? null
          //     // No children here — copied from OPP when user opens iFrame
          //   }));
          //   inserted++;
          //   console.log('Quote shell created: QuoteID=', opp.quoteId,
          //     'linked to OppID=', opp.Oppid);
          // }
          else {
            // First time — create Quote shell row
            const newQuoteID = cds.utils.uuid();
            await tx.run(INSERT.into(Opportunity).entries({
              ID: newQuoteID,
              Oppid: opp.Oppid,
              objectStatus: 'QUOTE',
              quoteId: opp.quoteId,
              status: opp.status ?? null,
              oppCategory: opp.oppCategory ?? null,
              salesPhase: opp.salesPhase ?? null
            }));
            inserted++;
            console.log('Quote shell created: QuoteID=', opp.quoteId, 'linked to OppID=', opp.Oppid);

            // ── Immediately copy OPP children + append payload rows ───────────
            const parentOppNew = await tx.run(
              SELECT.one.from(Opportunity).columns('ID')
                .where({ Oppid: opp.Oppid, objectStatus: 'OPP' })
            );

            // Copy ALL IP rows from OPP
            let newShellIPRows = [];
            if (parentOppNew) {
              const oppIP = await tx.run(
                SELECT.from(InvolvedParties).where({ opportunity_ID: parentOppNew.ID })
              );
              newShellIPRows = oppIP.map(r => ({
                ID: cds.utils.uuid(),
                buspartner: r.buspartner ?? null,
                role: r.role ?? null,
                opportunity_ID: newQuoteID
              }));
            }

            // Append NEW IP rows from payload (not already in OPP)
            const shellIPKeys = newShellIPRows.map(r => `${r.buspartner}|${r.role}`);
            const extraIPRows = (involvedParties || [])
              .filter(x => !shellIPKeys.includes(`${x?.buspartner}|${x?.role}`))
              .map(x => ({
                ID: cds.utils.uuid(),
                buspartner: x?.buspartner ?? null,
                role: x?.role ?? null,
                opportunity_ID: newQuoteID
              }));

            const allNewIPRows = [...newShellIPRows, ...extraIPRows];
            if (allNewIPRows.length > 0) {
              await insertInBatches(InvolvedParties, allNewIPRows);
              console.log(`New Quote IP: ${newShellIPRows.length} from OPP + ${extraIPRows.length} from payload`);
            }

            // Copy ALL PC rows from OPP
            let newShellPCRows = [];
            if (parentOppNew) {
              const oppPC = await tx.run(
                SELECT.from(ProviderContract).where({ opportunity_ID: parentOppNew.ID })
              );
              newShellPCRows = oppPC.map(r => ({
                ID: cds.utils.uuid(),
                product: r.product ?? null,
                portfolio: r.portfolio ?? null,
                fixedMWh: r.fixedMWh ?? null,
                fixedPrice: r.fixedPrice ?? null,
                portfolioprice: r.portfolioprice ?? null,
                subspercent: r.subspercent ?? null,
                startDate: r.startDate ?? null,
                endDate: r.endDate ?? null,
                enrollChk: r.enrollChk ?? null,
                exportDate: r.exportDate ?? null,
                exportChk: r.exportChk ?? null,
                netPremium: r.netPremium ?? null,
                opportunity_ID: newQuoteID
              }));
            }

            // Append NEW PC rows from payload (not already in OPP)
            const shellPCKeys = newShellPCRows.map(r => `${r.product}|${r.portfolio}|${r.startDate}`);
            const extraPCRows = (providerContracts || [])
              .filter(x => !shellPCKeys.includes(`${x?.product}|${x?.portfolio}`))
              .map(x => ({
                ID: cds.utils.uuid(),
                product: x?.product ?? null,
                portfolio: x?.portfolio ?? null,
                fixedMWh: x?.fixedMWh ?? null,
                fixedPrice: x?.fixedPrice ?? null,
                portfolioprice: x?.portfolioprice ?? null,
                subspercent: x?.subspercent ?? null,
                startDate: x?.startDate ?? null,
                endDate: x?.endDate ?? null,
                enrollChk: x?.enrollChk ?? null,
                exportDate: x?.exportDate ?? null,
                exportChk: x?.exportChk ?? null,
                opportunity_ID: newQuoteID
              }));

            const allNewPCRows = [...newShellPCRows, ...extraPCRows];
            if (allNewPCRows.length > 0) {
              await insertInBatches(ProviderContract, allNewPCRows);
              console.log(`New Quote PC: ${newShellPCRows.length} from OPP + ${extraPCRows.length} from payload`);
            }
          }

        } else {
          // ── OPP push — existing logic ─────────────────────────────────
          const existing = await tx.run(
            SELECT.one.from(Opportunity).columns('ID')
              .where({ Oppid: opp.Oppid, objectStatus: 'OPP' })
          );

          const oppCols = {
            quoteId: opp.quoteId ?? null,
            oppCategory: opp.oppCategory ?? null,
            status: opp.status ?? null,
            objectStatus: 'OPP',
            salesPhase: opp.salesPhase ?? null,
          };

          let oppID;
          if (existing) {
            await tx.run(UPDATE(Opportunity).set(oppCols).where({ ID: existing.ID }));
            oppID = existing.ID;
            updated++;
          } else {
            const newID = cds.utils.uuid();
            await tx.run(INSERT.into(Opportunity).entries({
              ID: newID, Oppid: opp.Oppid, ...oppCols
            }));
            oppID = newID;
            inserted++;
          }

          if (!oppID) return req.error(500, 'Failed to upsert Opportunity');

          // Replace children for OPP only
          await tx.run(DELETE.from(InvolvedParties).where({ opportunity_ID: oppID }));
          // await tx.run(DELETE.from(ProviderContract).where({ opportunity_ID: oppID }));

          const ipRows = (involvedParties || []).map(x => ({
            ID: cds.utils.uuid(),          // ← ADD THIS
            buspartner: x?.buspartner ?? null,
            role: x?.role ?? null,
            opportunity_ID: oppID
          }));

          // Snapshot existing exportChk/exportDate before deleting PC rows
          // Fetch ALL columns — avoid .columns() projection which can silently drop fields
          const existingPCSnap = await tx.run(
            SELECT.from(ProviderContract).where({ opportunity_ID: oppID })
          );
          const snapMap = {};
          existingPCSnap.forEach(r => { snapMap[`${r.product}|${r.portfolio}|${r.startDate}`] = r; });
          console.log('PC Snap rows fetched:', existingPCSnap.length);
          console.log('PC SnapMap:', JSON.stringify(snapMap, null, 2));
          await tx.run(DELETE.from(ProviderContract).where({ opportunity_ID: oppID }));

          // const pcRows = (providerContracts || []).map(x => ({
          //   ID: cds.utils.uuid(),          // ← ADD THIS
          //   product: x?.product ?? null,
          //   portfolio: x?.portfolio ?? null,
          //   fixedMWh: x?.fixedMWh ?? null,
          //   fixedPrice: x?.fixedPrice ?? null,
          //   portfolioprice: x?.portfolioprice ?? null,
          //   subspercent: x?.subspercent ?? null,
          //   startDate: x?.startDate ?? null,
          //   endDate: x?.endDate ?? null,
          //   enrollChk: x?.enrollChk ?? null,
          //   exportDate: x?.exportDate ?? null,
          //   exportChk: x?.exportChk ?? null,
          //   opportunity_ID: oppID
          // }));
          const pcRows = (providerContracts || []).map(x => {
            const snap = snapMap[`${x?.product}|${x?.portfolio}|${x?.startDate}`] || {};
            return {
              ID: cds.utils.uuid(),
              product: x?.product ?? null,
              portfolio: x?.portfolio ?? null,
              // ── C4C-owned fields: always take from payload ──
              startDate: x?.startDate ?? null,
              endDate: x?.endDate ?? null,
              // ── BTP-owned fields: restore from snap ──
              fixedMWh: snap.fixedMWh ?? null,
              fixedPrice: snap.fixedPrice ?? null,
              portfolioprice: snap.portfolioprice ?? null,
              subspercent: snap.subspercent ?? null,
              enrollChk: snap.enrollChk ?? null,
              exportChk: snap.exportChk ?? null,
              exportDate: snap.exportDate ?? null,
              // ── These were missing — add them ──
              recPrice: snap.recPrice ?? null,
              migpPercent: snap.migpPercent ?? null,
              recpercent: snap.recpercent ?? null,
              migpMWh: snap.migpMWh ?? null,
              recMWh: snap.recMWh ?? null,
              estUsage: snap.estUsage ?? null,
              netPremium: snap.netPremium ?? null,
              opportunity_ID: oppID
            };
          });


          childrenWritten += await insertInBatches(InvolvedParties, ipRows);
          childrenWritten += await insertInBatches(ProviderContract, pcRows);
        }

        return { status: 'SUCCESS', inserted, updated, childrenWritten };

      } catch (e) {
        req.error(500, e.message || e);
      }
    });

  }
};