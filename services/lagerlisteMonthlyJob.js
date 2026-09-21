const { validateValuation } = require('./lagerlisteAllocation');

function closingSlot(date) {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
        timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
    }).formatToParts(date).map(part => [part.type, part.value]));
    const lastDay = new Date(Date.UTC(Number(parts.year), Number(parts.month), 0)).getUTCDate();
    return {
        month: `${parts.year}-${parts.month}`,
        due: Number(parts.day) === lastDay && parts.hour === '23' && parts.minute === '59',
        lastDay: Number(parts.day) === lastDay
    };
}

// The period comes from the actual start, never from the completion date.
// Live calculations are explicitly labelled: this is not a SQL point-in-time snapshot.
async function runMonthlyClose({ now = () => new Date(), gohData, createService, writeBackup }) {
    const started = now();
    const slot = closingSlot(started);
    if (!slot.due) {
        if (slot.lastDay) throw new Error('Month-end job missed 23:59 Europe/Copenhagen; no retrospective closure was created.');
        return { status: 'not-due' };
    }
    const key = 'lagerliste_month_' + slot.month;
    const existing = await gohData.getAppState(key, { strict: true });
    if (existing) return { status: 'already-closed', month: slot.month };
    const service = createService(slot.month);
    const current = await service.getCurrent({ forceRefresh: true, forceAftercalc: true, valuationDate: started });
    const finished = now();
    validateValuation(current);
    if (!current.diverseStatus?.complete || current.diverseStatus.month !== slot.month) {
        throw new Error('Diverse is incomplete or belongs to another month. Closure not saved.');
    }
    const generated = new Date(current.generatedAt).getTime();
    if (!Number.isFinite(generated) || generated < started.getTime() || generated > finished.getTime() + 1000) {
        throw new Error('Calculation did not produce fresh data for this run.');
    }
    const payload = {
        month: slot.month, createdAt: finished.toISOString(), current, diverse: [],
        automaticClose: {
            timeZone: 'Europe/Copenhagen', scheduledLocalTime: '23:59',
            startedAt: started.toISOString(), completedAt: finished.toISOString(),
            valuationMode: 'live-during-calculation'
        }
    };
    if (!await gohData.setAppState(key, payload, { createOnly: true })) {
        const winner = await gohData.getAppState(key, { strict: true });
        if (winner) return { status: 'already-closed', month: slot.month };
        throw new Error('GOH did not confirm the closure.');
    }
    const stored = await gohData.getAppState(key, { strict: true });
    if (!stored || JSON.stringify(stored.payload) !== JSON.stringify(payload)) throw new Error('GOH read-back verification failed.');
    await writeBackup(slot.month, payload);
    return { status: 'saved', month: slot.month, startedAt: started.toISOString(), completedAt: finished.toISOString() };
}

function startClientMonthlyScheduler({ run, now = () => new Date(), setTimer = setInterval, onError = () => {}, onResult = () => {} }) {
    let attemptedMonth = null;
    let running = false;
    const tick = async () => {
        const slot = closingSlot(now());
        if (!slot.due || running || attemptedMonth === slot.month) return;
        attemptedMonth = slot.month;
        running = true;
        try { onResult(await run()); }
        catch (error) { onError(error); }
        finally { running = false; }
    };
    tick();
    const timer = setTimer(tick, 1000);
    if (timer && typeof timer.unref === 'function') timer.unref();
    return timer;
}

module.exports = { closingSlot, runMonthlyClose, startClientMonthlyScheduler };
