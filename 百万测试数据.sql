-- 1、创建表
CREATE TABLE `tb_user2` (
                            `id` bigint(20) NOT NULL  COMMENT '用户ID',
                            `user_name` varchar(100) DEFAULT NULL COMMENT '姓名',
                            `phone` varchar(15) DEFAULT NULL COMMENT '手机号',
                            `province` varchar(50) DEFAULT NULL COMMENT '省份',
                            `city` varchar(50) DEFAULT NULL COMMENT '城市',
                            `salary` int(10) DEFAULT NULL,
                            `hire_date` datetime DEFAULT NULL COMMENT '入职日期',
                            `dept_id` bigint(20) DEFAULT NULL COMMENT '部门编号',
                            `birthday` datetime DEFAULT NULL COMMENT '出生日期',
                            `photo` varchar(200) DEFAULT NULL COMMENT '照片路径',
                            `address` varchar(300) DEFAULT NULL COMMENT '现在住址'

) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- 2、创建存储过程
DELIMITER $$    -- 重新定义“;”分号
DROP PROCEDURE IF EXISTS test_insert $$   -- 如果有test_insert这个存储过程就删除
CREATE PROCEDURE test_insert()			  -- 创建存储过程

BEGIN
	DECLARE n int DEFAULT 1;				    -- 定义变量n=1
	SET AUTOCOMMIT=0;						    -- 取消自动提交

		while n <= 5000000 do
			INSERT INTO `tb_user2` VALUES ( n, CONCAT('测试', n), '13800000001', '北京市', '北京市', '11000', '2001-03-01 21:18:29', '1', '1981-03-02 00:00:00', '\\static\\user_photos\\1.jpg', '北京市西城区宣武大街1号院');
			SET n=n+1;
END while;
COMMIT;
END $$


-- 3、开始执行 插入500W数据大概需要200至300秒左右,执行时间较长
CALL test_insert();
