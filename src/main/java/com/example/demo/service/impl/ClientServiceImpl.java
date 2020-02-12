package com.example.demo.service.impl;

import com.example.demo.entity.Client;
import com.example.demo.repository.ClientRepository;
import com.example.demo.service.ClientService;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

/**
 * Service contenant les actions métiers liées aux clients.
 */
@Service
@Transactional
public class ClientServiceImpl implements ClientService {

    private ClientRepository clientRepository;

    public ClientServiceImpl(ClientRepository clientRepository) {
        this.clientRepository = clientRepository;
    }

    @Override
    public List<Client> findAllClients() {
        return clientRepository.findAll();
    }

    public Client findById(Long id) { return clientRepository.findById(id).get(); }
}
