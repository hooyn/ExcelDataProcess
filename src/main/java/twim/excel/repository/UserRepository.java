package twim.excel.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import twim.excel.entity.User;

public interface UserRepository extends JpaRepository<User, Long> {
}
