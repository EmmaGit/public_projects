var ViewModel = function () {
    var self = this;
    self.users = ko.observableArray();
    self.error = ko.observable();

    var usersUri = '/api/users/';
    var rolesUri = '/api/roles/';

    function ajaxHelper(uri, method, data) {
        self.error(''); // Clear error message
        return $.ajax({
            type: method,
            url: uri,
            dataType: 'json',
            contentType: 'application/json',
            data: data ? JSON.stringify(data) : null
        }).fail(function (jqXHR, textStatus, errorThrown) {
            self.error(errorThrown);
        });
    }

    function getAllUsers() {
        ajaxHelper(usersUri, 'GET').done(function (data) {
            self.users(data);
        });
    }

    // Fetch the initial data.
    getAllUsers();

    self.detail = ko.observable();

    self.getUserDetail = function (item) {
        ajaxHelper(usersUri + item.Id, 'GET').done(function (data) {
            self.detail(data);
        });
    }

    self.roles = ko.observableArray();
    
    self.newUser = {
        Role: ko.observable(),
        Username: ko.observable(),
        Email: ko.observable(),
        Picture: ko.observable(),
        Hostname: ko.observable()
    }

    function getRoles() {
        ajaxHelper(rolesUri, 'GET').done(function (data) {
            self.roles(data);
        });
    }

    self.addUser = function (formElement) {
        if (self.newUser.Role() == null)
        {
            var user = {
                Username: self.newUser.Username(),
                Email: self.newUser.Email(),
                Picture: self.newUser.Picture()
            }
        }
        else {
            var user = {
                RoleId: self.newUser.Role().Id,
                Username: self.newUser.Username(),
                Email: self.newUser.Email(),
                Picture: self.newUser.Picture(),
                Hostname: self.newUser.Hostname()

            }
        };

        ajaxHelper(usersUri, 'POST', user).done(function (item) {
            self.users.push(item);
        });
    }

    getRoles();
};

ko.applyBindings(new ViewModel());