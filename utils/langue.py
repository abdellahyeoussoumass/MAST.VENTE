langue_actuelle = ["fr"]

TRADUCTIONS = {
    "fr": {
        # Sidebar
        "app_name": "VentePro",
        "dashboard": "Dashboard",
        "produits": "Produits",
        "clients": "Clients",
        "ventes": "Ventes",
        "devis": "Devis",
        "bons_commande": "Bons de Commande",
        "factures": "Factures",
        "livraisons": "Livraisons",
        "guide": "Guide",
        "historique": "Historique",
        "profil": "Mon Profil",
        "theme": "Thème",
        "sauvegarde": "Sauvegarder",
        "deconnexion": "Déconnexion",
        # Login
        "connexion": "Se connecter",
        "nom_utilisateur": "Nom d'utilisateur",
        "mot_de_passe": "Mot de passe",
        "erreur_login": "❌ Nom ou mot de passe incorrect !",
        "connexion_cours": "Connexion en cours...",
        # Boutons communs
        "ajouter": "➕ Ajouter",
        "modifier": "✏️ Modifier",
        "supprimer": "🗑️ Supprimer",
        "rechercher": "🔍 Rechercher...",
        "importer": "📥 Importer Excel",
        "exporter": "📥 Exporter Excel",
        "imprimer": "🖨️ Imprimer",
        "envoyer": "📧 Envoyer",
        "sauvegarder": "💾 Sauvegarder",
        "annuler": "Annuler",
        # Messages
        "succes_ajout": "✅ Ajouté avec succès !",
        "succes_modif": "✅ Modifié avec succès !",
        "succes_suppr": "✅ Supprimé avec succès !",
        "confirmation": "Voulez-vous vraiment supprimer ?",
        "bienvenue": "👋 Bienvenue !",
        "selection_section": "Sélectionnez une section dans le menu.",
        # Backup
        "backup_succes": "Base de données sauvegardée !\n📁 Dossier : backups/",
    },
    "en": {
        # Sidebar
        "app_name": "VentePro",
        "dashboard": "Dashboard",
        "produits": "Products",
        "clients": "Clients",
        "ventes": "Sales",
        "devis": "Quotes",
        "bons_commande": "Purchase Orders",
        "factures": "Invoices",
        "livraisons": "Deliveries",
        "guide": "Guide",
        "historique": "History",
        "profil": "My Profile",
        "theme": "Theme",
        "sauvegarde": "Save",
        "deconnexion": "Logout",
        # Login
        "connexion": "Login",
        "nom_utilisateur": "Username",
        "mot_de_passe": "Password",
        "erreur_login": "❌ Incorrect username or password !",
        "connexion_cours": "Logging in...",
        # Boutons communs
        "ajouter": "➕ Add",
        "modifier": "✏️ Edit",
        "supprimer": "🗑️ Delete",
        "rechercher": "🔍 Search...",
        "importer": "📥 Import Excel",
        "exporter": "📥 Export Excel",
        "imprimer": "🖨️ Print",
        "envoyer": "📧 Send",
        "sauvegarder": "💾 Save",
        "annuler": "Cancel",
        # Messages
        "succes_ajout": "✅ Added successfully !",
        "succes_modif": "✅ Updated successfully !",
        "succes_suppr": "✅ Deleted successfully !",
        "confirmation": "Are you sure you want to delete ?",
        "bienvenue": "👋 Welcome !",
        "selection_section": "Select a section from the menu.",
        # Backup
        "backup_succes": "Database saved successfully !\n📁 Folder : backups/",
    },
    "ar": {
        # Sidebar
        "app_name": "ventePro",
        "dashboard": "لوحة التحكم",
        "produits": "المنتجات",
        "clients": "العملاء",
        "ventes": "المبيعات",
        "devis": "عروض الأسعار",
        "bons_commande": "أوامر الشراء",
        "factures": "الفواتير",
        "livraisons": "التوصيلات",
        "guide": "الدليل",
        "historique": "السجل",
        "profil": "ملفي الشخصي",
        "theme": "المظهر",
        "sauvegarde": "حفظ",
        "deconnexion": "تسجيل الخروج",
        # Login
        "connexion": "تسجيل الدخول",
        "nom_utilisateur": "اسم المستخدم",
        "mot_de_passe": "كلمة المرور",
        "erreur_login": "❌ اسم المستخدم أو كلمة المرور غير صحيحة !",
        "connexion_cours": "جاري تسجيل الدخول...",
        # Boutons communs
        "ajouter": "➕ إضافة",
        "modifier": "✏️ تعديل",
        "supprimer": "🗑️ حذف",
        "rechercher": "🔍 بحث...",
        "importer": "📥 استيراد Excel",
        "exporter": "📥 تصدير Excel",
        "imprimer": "🖨️ طباعة",
        "envoyer": "📧 إرسال",
        "sauvegarder": "💾 حفظ",
        "annuler": "إلغاء",
        # Messages
        "succes_ajout": "✅ تمت الإضافة بنجاح !",
        "succes_modif": "✅ تم التعديل بنجاح !",
        "succes_suppr": "✅ تم الحذف بنجاح !",
        "confirmation": "هل تريد حقاً الحذف ؟",
        "bienvenue": "👋 مرحباً !",
        "selection_section": "اختر قسماً من القائمة.",
        # Backup
        "backup_succes": "تم حفظ قاعدة البيانات بنجاح !\n📁 المجلد : backups/",
    }
}

def get_text(cle):
    return TRADUCTIONS[langue_actuelle[0]].get(cle, cle)

def changer_langue(nouvelle_langue):
    langue_actuelle[0] = nouvelle_langue

def get_langue():
    return langue_actuelle[0]