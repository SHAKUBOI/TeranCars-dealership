<?php
/**
 * Page panier
 * Affiche les véhicules ajoutés au panier et permet de passer commande
 */

// Définition du titre de la page
$pageTitle = "Panier";

// Inclusion des fichiers nécessaires
require_once '../config/config.php';

// Initialiser le panier s'il n'existe pas
if (!isset($_SESSION['panier'])) {
    $_SESSION['panier'] = [
        'items' => [],
        'sous_total' => 0,
        'tva' => 0,
        'total' => 0
    ];
}

// Fonction pour recalculer les totaux
function recalculerTotaux() {
    $sous_total = 0;
    foreach ($_SESSION['panier']['items'] as $item) {
        if ($item['type'] === 'achat') {
            $sous_total += $item['prix'] * $item['quantite'];
        } else {
            $sous_total += $item['prix'] * $item['quantite'] * $item['duree'];
        }
    }
    $_SESSION['panier']['sous_total'] = $sous_total;
    $_SESSION['panier']['tva'] = $sous_total * 0.2;
    $_SESSION['panier']['total'] = $sous_total + $_SESSION['panier']['tva'];
}

// Vérifier si la connexion à la base de données est établie
if (!isset($conn) || $conn->connect_error) {
    include '../includes/header.php';
    echo "<div class='container'><div class='alert alert-danger'>Erreur de connexion à la base de données. Veuillez réessayer plus tard.</div></div>";
    include '../includes/footer.php';
    exit();
}

// Vérifier si la table vehicules existe
$table_exists = false;
$check_table = $conn->query("SHOW TABLES LIKE 'vehicules'");
$table_exists = ($check_table && $check_table->num_rows > 0);

if (!$table_exists) {
    include '../includes/header.php';
    echo "<div class='container'><div class='alert alert-danger'>La table des véhicules n'existe pas. Veuillez contacter l'administrateur.</div></div>";
    include '../includes/footer.php';
    exit();
}

// Traitement des actions
if (isset($_GET['action'])) {
    $action = $_GET['action'];
    $error_message = null;
    
    // Ajouter un véhicule au panier
    if ($action === 'ajouter' && isset($_GET['id']) && isset($_GET['type'])) {
        $id_vehicule = (int)$_GET['id'];
        $type = $_GET['type']; // 'achat' ou 'location'
        
        // Vérifier si le véhicule existe
        $query = "SELECT * FROM vehicules WHERE id_vehicule = ?";
        $stmt = $conn->prepare($query);
        
        if ($stmt) {
            $stmt->bind_param("i", $id_vehicule);
            $stmt->execute();
            $result = $stmt->get_result();
            
            if ($result && $result->num_rows > 0) {
                $vehicule = $result->fetch_assoc();
                
                // Déterminer le prix en fonction du type
                $prix = 0;
                if ($type === 'achat') {
                    $prix = isset($vehicule['prix_achat']) ? $vehicule['prix_achat'] : $vehicule['prix'];
                } else {
                    $prix = isset($vehicule['prix_location']) ? $vehicule['prix_location'] : ($vehicule['prix'] * 0.02); // 2% du prix d'achat par défaut
                }
                
                // Vérifier si le véhicule est déjà dans le panier
                $item_exists = false;
                foreach ($_SESSION['panier']['items'] as $key => $item) {
                    if ($item['id_vehicule'] == $id_vehicule && $item['type'] == $type) {
                        // Incrémenter la quantité
                        $_SESSION['panier']['items'][$key]['quantite']++;
                        $item_exists = true;
                        break;
                    }
                }
                
                // Si le véhicule n'est pas dans le panier, l'ajouter
                if (!$item_exists) {
                    $_SESSION['panier']['items'][] = [
                        'id_vehicule' => $id_vehicule,
                        'marque' => $vehicule['marque'],
                        'modele' => $vehicule['modele'],
                        'prix' => $prix,
                        'type' => $type,
                        'quantite' => 1,
                        'duree' => ($type === 'location') ? 1 : 0 // Durée en jours pour la location
                    ];
                }
                
                // Recalculer les totaux
                recalculerTotaux();
                
                // Rediriger vers la page du panier
                header('Location: panier.php?status=added');
                exit;
            } else {
                $error_message = "Le véhicule demandé n'existe pas.";
            }
            $stmt->close();
        } else {
            $error_message = "Erreur lors de la préparation de la requête : " . $conn->error;
        }
    }
    
    // Supprimer un véhicule du panier
    elseif ($action === 'supprimer' && isset($_GET['index'])) {
        $index = (int)$_GET['index'];
        
        if (isset($_SESSION['panier']['items'][$index])) {
            // Supprimer l'élément
            array_splice($_SESSION['panier']['items'], $index, 1);
            
            // Recalculer le total
            recalculerTotaux();
            
            // Rediriger vers la page du panier
            $total = 0;
            foreach ($_SESSION['panier']['items'] as $item) {
                if ($item['type'] === 'achat') {
                    $total += $item['prix'] * $item['quantite'];
                } else {
                    $total += $item['prix'] * $item['quantite'] * $item['duree'];
                }
            }
            $_SESSION['panier']['total'] = $total;
            
            // Rediriger vers la page du panier
            header('Location: panier.php?status=removed');
            exit;
        }
    }
    
    // Mettre à jour la quantité
    elseif ($action === 'update_quantite' && isset($_GET['index']) && isset($_GET['quantite'])) {
        $index = (int)$_GET['index'];
        $quantite = max(1, (int)$_GET['quantite']); // Assurer un minimum de 1
        
        if (isset($_SESSION['panier']['items'][$index])) {
            $_SESSION['panier']['items'][$index]['quantite'] = $quantite;
            recalculerTotaux();
            
            // Retourner le nouveau total en JSON si c'est une requête AJAX
            if (isset($_SERVER['HTTP_X_REQUESTED_WITH']) && strtolower($_SERVER['HTTP_X_REQUESTED_WITH']) === 'xmlhttprequest') {
                header('Content-Type: application/json');
                echo json_encode([
                    'success' => true,
                    'sous_total' => $_SESSION['panier']['sous_total'],
                    'tva' => $_SESSION['panier']['tva'],
                    'total' => $_SESSION['panier']['total']
                ]);
                exit;
            }
            
            // Sinon, rediriger
            header('Location: panier.php?status=updated');
            exit;
        }
    }
    
    // Mettre à jour la durée de location
    elseif ($action === 'update_duree' && isset($_GET['index']) && isset($_GET['duree'])) {
        $index = (int)$_GET['index'];
        $duree = max(1, (int)$_GET['duree']); // Assurer un minimum de 1
        
        if (isset($_SESSION['panier']['items'][$index]) && $_SESSION['panier']['items'][$index]['type'] === 'location') {
            $_SESSION['panier']['items'][$index]['duree'] = $duree;
            recalculerTotaux();
            
            // Retourner le nouveau total en JSON si c'est une requête AJAX
            if (isset($_SERVER['HTTP_X_REQUESTED_WITH']) && strtolower($_SERVER['HTTP_X_REQUESTED_WITH']) === 'xmlhttprequest') {
                header('Content-Type: application/json');
                echo json_encode([
                    'success' => true,
                    'sous_total' => $_SESSION['panier']['sous_total'],
                    'tva' => $_SESSION['panier']['tva'],
                    'total' => $_SESSION['panier']['total']
                ]);
                exit;
            }
            
            // Sinon, rediriger
            header('Location: panier.php?status=updated');
            exit;
        }
    }
    
    // Vider le panier
    elseif ($action === 'vider') {
        $_SESSION['panier'] = [
            'items' => [],
            'total' => 0
        ];
        
        // Rediriger vers la page du panier
        header('Location: panier.php?status=emptied');
        exit;
    }
}

// Calculer les totaux
$sous_total = 0;
$tva = 0;
$total = 0;

if (isset($_SESSION['panier']['items']) && count($_SESSION['panier']['items']) > 0) {
    foreach ($_SESSION['panier']['items'] as $item) {
        if ($item['type'] === 'achat') {
            $sous_total += $item['prix'] * $item['quantite'];
        } else {
            $sous_total += $item['prix'] * $item['quantite'] * $item['duree'];
        }
    }
    
    // Calculer la TVA (20%)
    $tva = $sous_total * 0.2;
    $total = $sous_total + $tva;
    
    // Mettre à jour le total dans la session
    $_SESSION['panier']['total'] = $total;
}

// Inclure l'en-tête
include_once '../includes/header.php';
?>

<div class="container panier-page">
    <h1>Votre panier</h1>
    
    <?php if (isset($error_message)): ?>
        <div class="alert alert-danger"><?php echo $error_message; ?></div>
    <?php endif; ?>
    
    <?php if (isset($_GET['status'])): ?>
        <?php if ($_GET['status'] === 'added'): ?>
            <div class="alert alert-success">Le véhicule a été ajouté à votre panier.</div>
        <?php elseif ($_GET['status'] === 'removed'): ?>
            <div class="alert alert-info">Le véhicule a été retiré de votre panier.</div>
        <?php elseif ($_GET['status'] === 'updated'): ?>
            <div class="alert alert-info">Votre panier a été mis à jour.</div>
        <?php elseif ($_GET['status'] === 'emptied'): ?>
            <div class="alert alert-warning">Votre panier a été vidé.</div>
        <?php endif; ?>
    <?php endif; ?>
    
    <?php if (isset($_SESSION['panier']['items']) && count($_SESSION['panier']['items']) > 0): ?>
        <div class="panier-items">
            <?php foreach ($_SESSION['panier']['items'] as $index => $item): ?>
                <div class="panier-item">
                    <div class="item-image">
                        <?php
                        // Récupérer l'image du véhicule
                        $image_path = '../assets/images/default-car.jpg'; // Image par défaut
                        
                        // Vérifier si la table images_vehicules existe
                        $images_table_exists = false;
                        $check_images_table = $conn->query("SHOW TABLES LIKE 'images_vehicules'");
                        $images_table_exists = ($check_images_table && $check_images_table->num_rows > 0);
                        
                        if ($images_table_exists) {
                            $image_query = "SELECT chemin_image FROM images_vehicules WHERE id_vehicule = ? LIMIT 1";
                            $image_stmt = $conn->prepare($image_query);
                            
                            if ($image_stmt) {
                                $image_stmt->bind_param("i", $item['id_vehicule']);
                                $image_stmt->execute();
                                $image_result = $image_stmt->get_result();
                                
                                if ($image_result && $image_result->num_rows > 0) {
                                    $image_data = $image_result->fetch_assoc();
                                    $image_path = $image_data['chemin_image'];
                                }
                                $image_stmt->close();
                            }
                        }
                        ?>
                        <img src="<?php echo $image_path; ?>" alt="<?php echo htmlspecialchars($item['marque'] . ' ' . $item['modele']); ?>" class="img-fluid">
                    </div>
                    <div class="item-details">
                        <h3><?php echo htmlspecialchars($item['marque'] . ' ' . $item['modele']); ?></h3>
                        <p class="item-type"><?php echo ($item['type'] === 'achat') ? 'Achat' : 'Location'; ?></p>
                        <p class="item-price"><?php echo number_format($item['prix'], 2, ',', ' '); ?> €<?php echo ($item['type'] === 'location') ? '/jour' : ''; ?></p>
                        
                        <div class="item-actions">
                            <?php if ($item['type'] === 'achat'): ?>
                                <div class="quantity-control">
                                    <label for="quantite-<?php echo $index; ?>">Quantité:</label>
                                    <div class="input-group">
                                        <button class="btn btn-outline-secondary" type="button" onclick="updateQuantity(<?php echo $index; ?>, -1)">-</button>
                                        <input type="number" id="quantite-<?php echo $index; ?>" class="form-control" value="<?php echo $item['quantite']; ?>" min="1" onchange="updateQuantity(<?php echo $index; ?>, 0)">
                                        <button class="btn btn-outline-secondary" type="button" onclick="updateQuantity(<?php echo $index; ?>, 1)">+</button>
                                    </div>
                                </div>
                            <?php else: ?>
                                <div class="quantity-control">
                                    <label for="quantite-<?php echo $index; ?>">Quantité:</label>
                                    <div class="input-group">
                                        <button class="btn btn-outline-secondary" type="button" onclick="updateQuantity(<?php echo $index; ?>, -1)">-</button>
                                        <input type="number" id="quantite-<?php echo $index; ?>" class="form-control" value="<?php echo $item['quantite']; ?>" min="1" onchange="updateQuantity(<?php echo $index; ?>, 0)">
                                        <button class="btn btn-outline-secondary" type="button" onclick="updateQuantity(<?php echo $index; ?>, 1)">+</button>
                                    </div>
                                </div>
                                <div class="duration-control">
                                    <label for="duree-<?php echo $index; ?>">Durée (jours):</label>
                                    <div class="input-group">
                                        <button class="btn btn-outline-secondary" type="button" onclick="updateDuration(<?php echo $index; ?>, -1)">-</button>
                                        <input type="number" id="duree-<?php echo $index; ?>" class="form-control" value="<?php echo $item['duree']; ?>" min="1" onchange="updateDuration(<?php echo $index; ?>, 0)">
                                        <button class="btn btn-outline-secondary" type="button" onclick="updateDuration(<?php echo $index; ?>, 1)">+</button>
                                    </div>
                                </div>
                            <?php endif; ?>
                            
                            <div class="item-subtotal">
                                <p>Sous-total: <strong>
                                    <?php 
                                    $item_total = ($item['type'] === 'achat') 
                                        ? $item['prix'] * $item['quantite'] 
                                        : $item['prix'] * $item['quantite'] * $item['duree'];
                                    echo number_format($item_total, 2, ',', ' '); 
                                    ?> €</strong>
                                </p>
                            </div>
                            
                            <a href="panier.php?action=supprimer&index=<?php echo $index; ?>" class="btn btn-danger btn-sm">Supprimer</a>
                        </div>
                    </div>
                </div>
            <?php endforeach; ?>
            
            <div class="panier-summary">
                <div class="summary-details">
                    <div class="summary-row">
                        <span>Sous-total:</span>
                        <span><?php echo number_format($sous_total, 2, ',', ' '); ?> €</span>
                    </div>
                    <div class="summary-row">
                        <span>TVA (20%):</span>
                        <span><?php echo number_format($tva, 2, ',', ' '); ?> €</span>
                    </div>
                    <div class="summary-row total">
                        <span>Total:</span>
                        <span><?php echo number_format($total, 2, ',', ' '); ?> €</span>
                    </div>
                </div>
                
                <div class="panier-actions">
                    <a href="panier.php?action=vider" class="btn btn-warning">Vider le panier</a>
                    <a href="commande.php" class="btn btn-success">Passer commande</a>
                </div>
            </div>
        </div>
    <?php else: ?>
        <div class="panier-empty">
            <p>Votre panier est vide.</p>
            <a href="catalogue.php" class="btn btn-primary">Parcourir notre catalogue</a>
        </div>
    <?php endif; ?>
</div>

<script>
function updateQuantity(index, change) {
    const input = document.getElementById('quantite-' + index);
    let value = parseInt(input.value);
    
    if (change !== 0) {
        value += change;
        if (value < 1) value = 1;
        input.value = value;
    }
    
    // Faire une requête AJAX pour mettre à jour la quantité
    fetch('panier.php?action=update_quantite&index=' + index + '&quantite=' + value, {
        headers: {
            'X-Requested-With': 'XMLHttpRequest'
        }
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // Mettre à jour les totaux affichés
            document.querySelector('.summary-row .sous-total').textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(data.sous_total);
            document.querySelector('.summary-row .tva').textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(data.tva);
            document.querySelector('.summary-row.total .value').textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(data.total);
            
            // Mettre à jour le sous-total de l'article
            const itemTotal = document.querySelector('#item-total-' + index);
            if (itemTotal) {
                const prix = parseFloat(itemTotal.getAttribute('data-prix'));
                const duree = parseInt(itemTotal.getAttribute('data-duree') || 1);
                itemTotal.textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(prix * value * duree);
            }
        }
    })
    .catch(error => {
        console.error('Erreur lors de la mise à jour:', error);
        window.location.href = 'panier.php?action=update_quantite&index=' + index + '&quantite=' + value;
    });
}

function updateDuration(index, change) {
    const input = document.getElementById('duree-' + index);
    let value = parseInt(input.value);
    
    if (change !== 0) {
        value += change;
        if (value < 1) value = 1;
        input.value = value;
    }
    
    // Faire une requête AJAX pour mettre à jour la durée
    fetch('panier.php?action=update_duree&index=' + index + '&duree=' + value, {
        headers: {
            'X-Requested-With': 'XMLHttpRequest'
        }
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // Mettre à jour les totaux affichés
            document.querySelector('.summary-row .sous-total').textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(data.sous_total);
            document.querySelector('.summary-row .tva').textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(data.tva);
            document.querySelector('.summary-row.total .value').textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(data.total);
            
            // Mettre à jour le sous-total de l'article
            const itemTotal = document.querySelector('#item-total-' + index);
            if (itemTotal) {
                const prix = parseFloat(itemTotal.getAttribute('data-prix'));
                const quantite = parseInt(document.getElementById('quantite-' + index).value);
                itemTotal.textContent = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(prix * quantite * value);
            }
        }
    })
    .catch(error => {
        console.error('Erreur lors de la mise à jour:', error);
        window.location.href = 'panier.php?action=update_duree&index=' + index + '&duree=' + value;
    });
}

// Mettre à jour le compteur du panier dans le header
function updateCartCount() {
    const cartCount = document.getElementById('cart-count');
    if (cartCount) {
        cartCount.textContent = '<?php echo count($_SESSION['panier']['items']); ?>';
    }
}

// Appeler updateCartCount au chargement de la page
document.addEventListener('DOMContentLoaded', updateCartCount);
</script>

<?php include_once '../includes/footer.php'; ?>
