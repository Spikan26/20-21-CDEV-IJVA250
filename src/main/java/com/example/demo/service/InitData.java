package com.example.demo.service;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.ApplicationListener;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import javax.persistence.EntityManager;
import java.time.LocalDate;

/**
 * Classe permettant d'insérer des données dans l'application.
 */
@Service
@Transactional
public class InitData implements ApplicationListener<ApplicationReadyEvent> {

    private EntityManager em;

    public InitData(EntityManager em) {
        this.em = em;
    }

    @Override
    public void onApplicationEvent(ApplicationReadyEvent applicationReadyEvent) {
        insertTestData();
    }

    private void insertTestData() {
        Client client1 = newClient("Collet", "Adrienne", LocalDate.of(1979, 11, 6));
        em.persist(client1);

        Client client2 = newClient("Brunet", "Valérie", LocalDate.of(1997, 6, 11));
        em.persist(client2);

        em.persist(newClient("Hardy", "Thierry-Eugène", LocalDate.of(1996, 5, 14)));
        em.persist(newClient("ABSCHEN", "Jean", LocalDate.of(1996, 5, 14)));
        em.persist(newClient("ADAMO", "Stéphane", LocalDate.of(1986, 5, 14)));
        em.persist(newClient("AMELLAL", "Viviane", LocalDate.of(1996, 5, 14)));
        em.persist(newClient("ANGONIN", "Jean-Pierre", LocalDate.of(1991, 5, 14)));
        em.persist(newClient("AZOURA", "Marie-France", LocalDate.of(1996, 5, 14)));
        em.persist(newClient("AZRIA", "Maryse", LocalDate.of(1996, 1, 14)));
        em.persist(newClient("BACH", "Sylvie", LocalDate.of(1996, 5, 14)));
        em.persist(newClient("BARNAUD", "Janine", LocalDate.of(1993, 5, 14)));
        em.persist(newClient("BENSIMHON", "Pascal", LocalDate.of(1996, 5, 24)));
        em.persist(newClient("BERTRAND", "Roger", LocalDate.of(1996, 5, 14)));
        em.persist(newClient("BIDAULT", "Marie-Reine", LocalDate.of(1991, 5, 14)));
        em.persist(newClient("BINET", "Emmanuel", LocalDate.of(1990, 5, 18)));
        em.persist(newClient("BLANC", "Giséle", LocalDate.of(1999, 1, 14)));
        em.persist(newClient("BLANCHOT", "Guy", LocalDate.of(1996, 1, 17)));
        em.persist(newClient("BOUCHET", "Micheline", LocalDate.of(1996, 5, 14)));
        em.persist(newClient("BOUDART", "Orianne", LocalDate.of(1993, 5, 15)));
        em.persist(newClient("BOULLICAUD", "Paul", LocalDate.of(1995, 5, 14)));

        Article a1 = new Article();
        a1.setLibelle("Chargeurs de téléphones Portables");
        a1.setPrix(22.98);
        em.persist(a1);

        Article a2 = new Article();
        a2.setLibelle("Playmobil Hydravion de Police");
        a2.setPrix(14.39);
        em.persist(a2);

        Article a3 = new Article();
        a3.setLibelle("Distributeur de croquettes pour chien");
        a3.setPrix(12.99);
        em.persist(a3);

        Facture f1 = new Facture();
        f1.setClient(client1);
        em.persist(f1);
        em.persist(newLigneFacture(f1, a1, 2));
        em.persist(newLigneFacture(f1, a2, 1));

        Facture f2 = new Facture();
        f2.setClient(client2);
        em.persist(f2);
        em.persist(newLigneFacture(f2, a1, 10));

        Facture f3 = new Facture();
        f3.setClient(client1);
        em.persist(f3);
        em.persist(newLigneFacture(f3, a3, 1));

    }


    private LigneFacture newLigneFacture(Facture f, Article a1, int quantite) {
        LigneFacture lf1 = new LigneFacture();
        lf1.setArticle(a1);
        lf1.setQuantite(quantite);
        lf1.setFacture(f);
        return lf1;
    }


    private Client newClient(String nom, String prenom, LocalDate dateNaissance) {
        Client client = new Client();
        client.setNom(nom);
        client.setPrenom(prenom);
        client.setDateNaissance(dateNaissance);
        return client;
    }

}
