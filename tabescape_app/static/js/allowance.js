/* Browser-local offer timing, never a gameplay or paid entitlement gate. */
window.DailyAllowance = {
    limit: 1,
    day(date = new Date()) {
        return [date.getFullYear(), date.getMonth() + 1, date.getDate()].join("-");
    },
    read(storage, date = new Date()) {
        const day = this.day(date);
        try {
            const value = JSON.parse(storage.get("te_daily_free_stories_v2", "{}"));
            if (value.day === day && Array.isArray(value.completed)) {
                return {day, completed: [...new Set(value.completed.filter(id => typeof id === "string"))]};
            }
        } catch { /* Invalid settings start a new allowance. */ }
        return {day, completed: []};
    },
    remaining(storage, date = new Date()) {
        return Math.max(0, this.limit - this.read(storage, date).completed.length);
    },
    complete(storage, id, date = new Date()) {
        const record = this.read(storage, date);
        if (!record.completed.includes(id)) record.completed.push(id);
        storage.set("te_daily_free_stories_v2", JSON.stringify(record));
        return this.remaining(storage, date);
    }
};
