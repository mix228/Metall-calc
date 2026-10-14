/*
Переключатель тёмной темы.
Хранит выбор в localStorage, применяет атрибут data-theme до отрисовки.
*/
(function () {
    var KEY = "calc_theme";

    function apply(theme) {
        document.documentElement.setAttribute("data-theme", theme);
    }

    var saved = "light";
    try {
        saved = localStorage.getItem(KEY) || "light";
    } catch (e) {
        saved = "light";
    }
    apply(saved);

    document.addEventListener("DOMContentLoaded", function () {
        var btn = document.getElementById("theme-toggle");
        if (!btn) return;

        function label(theme) {
            return theme === "dark" ? "Светлая тема" : "Тёмная тема";
        }

        btn.textContent = label(saved);

        btn.addEventListener("click", function () {
            var cur = document.documentElement.getAttribute("data-theme") === "dark" ? "dark" : "light";
            var next = cur === "dark" ? "light" : "dark";
            apply(next);
            try {
                localStorage.setItem(KEY, next);
            } catch (e) {}
            btn.textContent = label(next);
        });
    });
})();