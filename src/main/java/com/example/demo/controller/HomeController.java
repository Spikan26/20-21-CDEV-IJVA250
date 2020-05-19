package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.repository.ArticleRepository;
import com.example.demo.service.impl.ClientServiceImpl;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.servlet.ModelAndView;

import org.apache.poi.ss.usermodel.CellStyle;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Optional;

/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController {

    private ArticleService articleService;
    private ClientServiceImpl clientServiceImpl;
    private FactureService factureService;

    public HomeController(ArticleService articleService, ClientServiceImpl clientService, FactureService factureService) {
        this.articleService = articleService;
        this.clientServiceImpl = clientService;
        this.factureService = factureService;
    }

    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Article> articles = articleService.findAll();
        modelAndView.addObject("articles", articles);

        List<Client> toto = clientServiceImpl.findAllClients();
        modelAndView.addObject("clients", toto);

        List<Facture> factures = factureService.findAllFactures();
        modelAndView.addObject("factures", factures);

        return modelAndView;
    }

    @GetMapping("/articles/csv")
    public void articlesCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition","attachment; filename=\"export-articles.csv\"");
        PrintWriter writer = response.getWriter();

        writer.println("Libelle;Prix");
        List<Article> articles = articleService.findAll();

        for (Article article : articles){
            String line = article.getLibelle() + ";" + article.getPrix();
            writer.println(line);
        }
    }

    @GetMapping("/articles/xlsx")
    public void articlesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition","attachment; filename=\"export-articles.xlsx\"");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Articles");

        //Font style for header
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setColor(IndexedColors.PINK.getIndex());

        //Header with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);


        Row headerRow = sheet.createRow(0);
        Cell cellLibelle = headerRow.createCell(0);
        Cell cellPrix = headerRow.createCell(1);

        cellLibelle.setCellValue("Libelle");
        cellLibelle.setCellStyle(headerCellStyle);
        cellPrix.setCellValue("Prix");
        cellPrix.setCellStyle(headerCellStyle);


        Integer rowNum = 1;
        List<Article> articles = articleService.findAll();

        for(Article article : articles){
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(article.getLibelle());

            row.createCell(1)
                    .setCellValue(article.getPrix());

        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }



    @GetMapping("/clients/csv")
    public void clientCSV(HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();

        writer.println("Nom;Prenom;Age");
        List<Client> clients = clientServiceImpl.findAllClients();

        for (Client client : clients){
            LocalDate now = LocalDate.now();
            Integer age = client.getDateNaissance().until(now).getYears();


            String line = client.getNom() + ";" + client.getPrenom() + ";" + age;
            writer.println(line);
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition","attachment; filename=\"export-clients.xlsx\"");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");

        //Font style for header
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setColor(IndexedColors.PINK.getIndex());

        //Header with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);


        Row headerRow = sheet.createRow(0);
        Cell cellNom = headerRow.createCell(0);
        Cell cellPrenom = headerRow.createCell(1);
        Cell cellAge = headerRow.createCell(2);

        cellNom.setCellValue("Nom");
        cellNom.setCellStyle(headerCellStyle);
        cellPrenom.setCellValue("Prénom");
        cellPrenom.setCellStyle(headerCellStyle);
        cellAge.setCellValue("Age");
        cellAge.setCellStyle(headerCellStyle);

        Integer rowNum = 1;
        List<Client> clients = clientServiceImpl.findAllClients();

        for(Client client : clients){
            Row row = sheet.createRow(rowNum++);

            LocalDate now = LocalDate.now();
            Integer age = client.getDateNaissance().until(now).getYears();

            row.createCell(0)
                    .setCellValue(client.getNom());


            row.createCell(1)
                    .setCellValue(client.getPrenom());

            row.createCell(2)
                    .setCellValue(age);
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }


    @GetMapping("/clients/{id}/factures/xlsx")
    public void facturesClientXLSX(HttpServletRequest request, HttpServletResponse response, @PathVariable("id") Long id ) throws IOException{
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition","attachment; filename=\"export-factures-clients.xlsx\"");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");

        Client client = clientServiceImpl.findById(id);
        List<Facture> factures = factureService.findAllFactures();


        //Font style for header
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setColor(IndexedColors.PINK.getIndex());

        //Header with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);


        Row headerRow = sheet.createRow(0);
        Cell cellNom = headerRow.createCell(0);
        Cell cellPrenom = headerRow.createCell(1);
        Cell cellAge = headerRow.createCell(2);

        cellNom.setCellValue("Nom");
        cellNom.setCellStyle(headerCellStyle);
        cellPrenom.setCellValue("Prénom");
        cellPrenom.setCellStyle(headerCellStyle);
        cellAge.setCellValue("Age");
        cellAge.setCellStyle(headerCellStyle);

        for(Facture facture : factures){
            if(facture.getClient() == client){
                Sheet factsheet = workbook.createSheet(facture.getId().toString());

                Row factheaderRow = factsheet.createRow(0);
                Cell cellLigneFactures = factheaderRow.createCell(0);

                cellLigneFactures.setCellValue("LigneFacture");
                cellLigneFactures.setCellStyle(headerCellStyle);

                Row row = factsheet.createRow(1);

                row.createCell(0)
                        .setCellValue(facture.getLigneFactures().size());
            }
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
