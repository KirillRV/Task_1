package com.calculator.dao;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

public class UserDAO {

    // In-memory storage of users: username -> password
    private static final Map<String, String> users = new HashMap<>();

    // Register new user if username doesn't exist
    public static boolean register(String username, String password) {
        if (users.containsKey(username)) {
            return false;
        }
        users.put(username, password);
        return true;
    }

    // Validate user credentials
    public static boolean validate(String username, String password) {
        return users.containsKey(username) && users.get(username).equals(password);
    }

    // Return set of all registered usernames
    public static Set<String> getAllUsers() {
        return users.keySet();
    }


}
