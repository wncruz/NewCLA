<!-- CONTEUDO DA PAGINA TERMINA AQUI -->
    </div><!-- fecha content-wrapper -->

    <script>
        var currentOpenDropdown = null;
        function toggleDropdown(dropdownId) {
            var dropdown = document.getElementById(dropdownId);
            var allDropdowns = document.querySelectorAll('.dropdown-menu');
            var allLinks = document.querySelectorAll('.nav-link');
            for (var i = 0; i < allDropdowns.length; i++) {
                if (allDropdowns[i].id !== dropdownId) allDropdowns[i].classList.remove('active');
            }
            for (var i = 0; i < allLinks.length; i++) allLinks[i].classList.remove('active');
            if (dropdown) {
                var isActive = dropdown.classList.contains('active');
                if (!isActive) {
                    dropdown.classList.add('active');
                    var link = dropdown.previousElementSibling;
                    if (link) link.classList.add('active');
                    currentOpenDropdown = dropdown;
                } else {
                    dropdown.classList.remove('active');
                    currentOpenDropdown = null;
                }
            }
        }
        document.addEventListener('click', function(e) {
            var target = e.target;
            if (!target.classList.contains('nav-link') && !target.classList.contains('dropdown-item')) {
                var allDropdowns = document.querySelectorAll('.dropdown-menu');
                var allLinks = document.querySelectorAll('.nav-link');
                for (var i = 0; i < allDropdowns.length; i++) allDropdowns[i].classList.remove('active');
                for (var i = 0; i < allLinks.length; i++) allLinks[i].classList.remove('active');
                currentOpenDropdown = null;
            }
        });
        var mobileBtn = document.getElementById('mobileMenuBtn');
        var mainMenu = document.getElementById('mainMenu');
        if (mobileBtn) {
            mobileBtn.addEventListener('click', function(e) {
                e.stopPropagation();
                mainMenu.classList.toggle('mobile-open');
            });
        }
        var navLinks = document.querySelectorAll('.nav-link');
        for (var i = 0; i < navLinks.length; i++) {
            navLinks[i].addEventListener('click', function(e) { e.stopPropagation(); });
        }
    </script>
</body>
</html>