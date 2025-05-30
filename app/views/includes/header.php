<header class="header">
    <nav class="navbar navbar-expand-lg navbar-light bg-light">
        <div class="container">
            <a class="navbar-brand fw-bold" href="<?= url('/') ?>">TeranCar</a>

            <!-- Menu burger pour mobile -->
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav"
                aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
                <span class="navbar-toggler-icon"></span>
            </button>

            <!-- Liens de navigation -->
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item"><a class="nav-link" href="<?= url('catalogue') ?>">Catalogue</a></li>
                    <li class="nav-item"><a class="nav-link" href="<?= url('about') ?>">À propos</a></li>
                    <li class="nav-item"><a class="nav-link" href="<?= url('contact') ?>">Contact</a></li>
                    <li class="nav-item">
                        <a class="btn btn-outline-primary btn-sm px-4" href="<?= url('auth/login') ?>">Se connecter</a>
                    </li>
                </ul>
            </div>
        </div>
    </nav>
</header>


<!-- Bootstrap CSS -->
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<!-- Bootstrap Icons -->
<link href="https://cdn.jsdelivr.net/npm/bootstrap-icons/font/bootstrap-icons.css" rel="stylesheet">

<!-- Bootstrap JS pour le menu burger -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>