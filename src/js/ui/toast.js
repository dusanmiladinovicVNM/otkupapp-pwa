function showToast(msg, type = 'info') {
    const toast = byId('toast');
    setText(toast, msg);
    toast.className = 'toast show ' + type;
    setTimeout(() => { toast.className = 'toast'; }, 3000);
}
