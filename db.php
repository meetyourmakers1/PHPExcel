<?php

class db
{
    public $conn = null;

    public function __construct($config)
    {
        $this->conn = mysqli_connect($config['host'],$config['username'],$config['password'],$config['database']) or die(mysqli_error($this->conn));
        mysqli_set_charset($this->conn,$config['charset']) or die(mysqli_error($this->conn));
    }

    public function getData($sql)
    {
        $result = mysqli_query($this->conn,$sql) or die(mysqli_error($this->conn));
        $data = array();
        while(($row = mysqli_fetch_assoc($result)) != false)
        {
            $data[] = $row;
        }
        return $data;
    }

    public function getDataByGrade($grade)
    {
        $sql = "select * from student where grade = ".$grade;
        $data = $this->getData($sql);
        return $data;
    }

    public function getAllGrade()
    {
        $sql = "select distinct(grade) from student".' order by grade asc';
        $data = $this->getData($sql);
        return $data;
    }

    public function getAllClass($grade)
    {
        $sql = 'select distinct(class) from student where grade = '.$grade.' order by class asc';
        $data = $this->getData($sql);
        return $data;
    }

    public function getAllStudent($grade,$class)
    {
        $sql = 'select name,score from student where grade = '.$grade.' and class = '.$class.' order by score desc';
        $data = $this->getData($sql);
        return $data;
    }
}