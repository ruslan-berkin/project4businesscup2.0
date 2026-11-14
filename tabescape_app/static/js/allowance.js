/* Browser-local demo allowance. Paid entitlements will be verified by the server. */
window.DailyAllowance = {
    limit: 3,
    day(date = new Date()) {
        return [date.getFullYear(), date.getMonth() + 1, date.getDate()].join("-");
    },
    read(storage, date = new Date()) {
        const day = this.day(date);
        try {
            const value = JSON.parse(storage.get("te_daily_stories", "{}"));
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
        storage.set("te_daily_stories", JSON.stringify(record));
        return this.remaining(storage, date);
    }
};
