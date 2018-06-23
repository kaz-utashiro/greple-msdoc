requires 'perl', '5.008001';

requires 'App::Greple';
requires 'List::Util', '1.45';

on 'test' => sub {
    requires 'Test::More', '0.98';
};

